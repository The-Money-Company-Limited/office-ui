// The office-ui build box on Cloudflare: one container (4 vCPU, 12 GiB, 20 GB disk) owned by a
// Durable Object, driven by tools/cloudflare/r over HTTP. Every request needs the RUNNER_TOKEN
// secret. The box stops 30 minutes after its last request.
import { DurableObject } from "cloudflare:workers";

interface Env {
  BUILDER: DurableObjectNamespace<Builder>;
  RUNNER_TOKEN: string;
}

const IDLE_MS = 30 * 60_000;
const encoder = new TextEncoder();

export class Builder extends DurableObject<Env> {
  /** The running box, started from the `base` image when it isn't running. */
  private async box() {
    const container = this.ctx.container;
    if (!container) throw new Error("No container is configured");
    if (!container.running) {
      container.start({
        image: container.images.base,
        instance: "standard-4",
        enableInternet: true,
        entrypoint: ["sleep", "infinity"],
      });
      // start() returns before the box can run commands.
      for (let attempt = 0; ; attempt++) {
        try {
          if ((await (await container.exec(["true"])).exitCode) === 0) break;
        } catch (error) {
          if (attempt >= 60) throw error;
        }
        await new Promise((resolve) => setTimeout(resolve, 1000));
      }
    }
    container.setInactivityTimeout(IDLE_MS);
    return container;
  }

  /**
   * Runs `bash -lc command` in `cwd` with `stdin` as its input. With `exitLine`, the output (stdout
   * and stderr together) ends with `__EXIT__ <code>` on its own line; without, it is the bare
   * stdout, for downloads.
   */
  async run(command: string, cwd: string, stdin: ReadableStream | null, exitLine: boolean): Promise<ReadableStream> {
    const box = await this.box();
    const process = await box.exec(["bash", "-lc", command], {
      cwd,
      stdin: stdin ?? undefined,
      stderr: exitLine ? "combined" : "ignore",
    });
    const { readable, writable } = new TransformStream<Uint8Array, Uint8Array>();
    void (async () => {
      if (process.stdout) await process.stdout.pipeTo(writable, { preventClose: true });
      const writer = writable.getWriter();
      if (exitLine) await writer.write(encoder.encode(`\n__EXIT__ ${await process.exitCode}\n`));
      await writer.close();
    })();
    return readable;
  }

  async down(): Promise<void> {
    if (this.ctx.container?.running) await this.ctx.container.destroy();
  }
}

export default {
  async fetch(request: Request, env: Env): Promise<Response> {
    if (!env.RUNNER_TOKEN || request.headers.get("authorization") !== `Bearer ${env.RUNNER_TOKEN}`) {
      return new Response("unauthorized\n", { status: 401 });
    }
    const url = new URL(request.url);
    const box = env.BUILDER.getByName("office-builds");
    const cwd = url.searchParams.get("cwd") ?? "/work";
    const text = (body: ReadableStream) => new Response(body, { headers: { "content-type": "text/plain; charset=utf-8" } });
    switch (`${request.method} ${url.pathname}`) {
      // A command; its output streams back, ending with the exit code line.
      case "POST /run": {
        const { command } = (await request.json()) as { command: string };
        return text(await box.run(command, cwd, null, true));
      }
      // Bytes from the request body into a file (a chunk of an upload).
      case "PUT /file": {
        const path = url.searchParams.get("path");
        if (!path) return new Response("path is required\n", { status: 400 });
        const quoted = `'${path.replaceAll("'", "'\\''")}'`;
        return text(await box.run(`mkdir -p "$(dirname ${quoted})" && cat > ${quoted}`, "/", request.body, true));
      }
      // A command's bare stdout, such as `tar czf - dist`.
      case "POST /stdout": {
        const { command } = (await request.json()) as { command: string };
        return new Response(await box.run(command, cwd, null, false));
      }
      case "POST /down":
        await box.down();
        return new Response("down\n");
      default:
        return new Response("not found\n", { status: 404 });
    }
  },
};
