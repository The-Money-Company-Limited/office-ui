// The office-ui build box on Cloudflare: one container (4 vCPU, 12 GiB, 20 GB disk) owned by a
// Durable Object, driven by tools/cloudflare/r over HTTP. Every request needs the RUNNER_TOKEN
// secret. Everyone with the token shares the one box, which stops 30 minutes after its last
// request; its files go with it.
import { DurableObject } from "cloudflare:workers";

interface Env {
  BUILDER: DurableObjectNamespace<Builder>;
  RUNNER_TOKEN?: string;
}

const IDLE_MS = 30 * 60_000;
const START_SECONDS = 60;
const encoder = new TextEncoder();

export class Builder extends DurableObject<Env> {
  /** The running box, started from the `base` image (which sleeps) when it isn't running. */
  private async box() {
    const container = this.ctx.container;
    if (!container) throw new Error("No container is configured");
    if (!container.running) {
      const image = container.images.base;
      if (!image) throw new Error(`No base image; the deploy gave ${JSON.stringify(container.images)}`);
      container.start({ image, instance: "standard-4", enableInternet: true });
      // start() returns before the box can run commands.
      for (let second = 0; ; second++) {
        const ready = await container.exec(["true"]).then((p) => p.exitCode).catch(() => null);
        if (ready === 0) break;
        if (second >= START_SECONDS) throw new Error(`the box did not start within ${START_SECONDS} s; try again`);
        await new Promise((resolve) => setTimeout(resolve, 1000));
      }
    }
    container.setInactivityTimeout(IDLE_MS);
    return container;
  }

  /**
   * Runs `bash -lc command` in `cwd` with `stdin` as its input. Its output streams back: stdout and
   * stderr together, ending with `__EXIT__ <code>` on its own line ("exit"), or the bare stdout,
   * for downloads ("raw").
   */
  async run(command: string, cwd: string, stdin: ReadableStream | null, output: "exit" | "raw"): Promise<ReadableStream> {
    const box = await this.box();
    // The shell changes to cwd itself: given a directory that isn't there, exec reports `bash` as
    // not found.
    const enter = '[ -d "$1" ] && cd -- "$1" || { echo "$1 is not on the box. Run r sync first: a box that stopped starts empty." >&2; exit 1; }; exec bash -lc "$2"';
    const process = await box.exec(["bash", "-c", enter, "bash", cwd, command], {
      cwd: "/",
      stdin: stdin ?? undefined,
      stderr: output === "exit" ? "combined" : "ignore",
    });
    const { readable, writable } = new TransformStream<Uint8Array, Uint8Array>();
    void (async () => {
      let code: number | null = null;
      try {
        if (process.stdout) await process.stdout.pipeTo(writable, { preventClose: true });
        code = await process.exitCode;
      } catch {
        // The finally below still ends the stream, reporting the failure as exit code 1.
      } finally {
        // A command killed (out of memory) or a broken pipe still ends the stream, as a failure.
        const writer = writable.getWriter();
        if (output === "exit") await writer.write(encoder.encode(`\n__EXIT__ ${code ?? 1}\n`)).catch(() => {});
        await writer.close().catch(() => {});
      }
    })();
    return readable;
  }

  async down(): Promise<void> {
    if (this.ctx.container?.running) await this.ctx.container.destroy();
  }
}

export default {
  async fetch(request: Request, env: Env): Promise<Response> {
    if (!env.RUNNER_TOKEN) {
      return new Response("RUNNER_TOKEN is not set on the office-builds Worker (see tools/cloudflare/r for setup)\n", { status: 500 });
    }
    if (request.headers.get("authorization") !== `Bearer ${env.RUNNER_TOKEN}`) {
      return new Response("unauthorized\n", { status: 401 });
    }
    const url = new URL(request.url);
    const box = env.BUILDER.getByName("office-builds");
    try {
      switch (`${request.method} ${url.pathname}`) {
        // ?command=…&cwd=…&output=exit|raw, with the request body (if any) as the command's input.
        case "POST /run": {
          const command = url.searchParams.get("command");
          if (!command) return new Response("command is required\n", { status: 400 });
          const output = url.searchParams.get("output") === "raw" ? "raw" : "exit";
          const body = await box.run(command, url.searchParams.get("cwd") ?? "/work", request.body, output);
          return new Response(body, { headers: { "content-type": "text/plain; charset=utf-8" } });
        }
        case "POST /down":
          await box.down();
          return new Response("down\n");
        default:
          return new Response("not found: use POST /run or POST /down\n", { status: 404 });
      }
    } catch (error) {
      return new Response(`${error instanceof Error ? error.message : error}\n`, { status: 503 });
    }
  },
};
