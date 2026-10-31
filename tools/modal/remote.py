# Runs builds and tests for office-ui (and repos such as office-env) on a Modal
# sandbox instead of the laptop. A <repo> is a path relative to ~/Desktop/Code
# (e.g. office-ui, or office-ui/engines/duke-sheets); it is uploaded to /work/<repo>, keeping
# the sandbox's node_modules and target dirs between runs.
#
#   uv run --with modal python remote.py up                 # create or reuse the sandbox
#   uv run --with modal python remote.py sync <repo>...     # upload tracked + new files
#   uv run --with modal python remote.py run <repo> '<cmd>' # run in /work/<repo>
#   uv run --with modal python remote.py pull <repo> <path> # copy a dir back (e.g. dist)
#   uv run --with modal python remote.py down
import io
import os
import subprocess
import sys
import tarfile

import modal

CODE = os.path.expanduser("~/Desktop/Code")
APP = "office-builds"
NAME = "office-builds-box"

image = (
    modal.Image.debian_slim(python_version="3.12")
    .apt_install("curl", "git", "build-essential", "pkg-config", "ca-certificates", "xz-utils", "unzip")
    .run_commands(
        "curl -fsSL https://deb.nodesource.com/setup_22.x | bash - && apt-get install -y nodejs",
        "npm install -g pnpm@9.15.0",
        "curl https://sh.rustup.rs -sSf | sh -s -- -y --profile minimal --target wasm32-unknown-unknown",
        "curl -fsSL https://github.com/rustwasm/wasm-pack/releases/download/v0.13.1/wasm-pack-v0.13.1-x86_64-unknown-linux-musl.tar.gz"
        " | tar xz -C /usr/local/bin --strip-components=1 wasm-pack-v0.13.1-x86_64-unknown-linux-musl/wasm-pack",
    )
    .env({"PATH": "/root/.cargo/bin:/usr/local/sbin:/usr/local/bin:/usr/sbin:/usr/bin:/sbin:/bin", "CI": "1"})
)


def sandbox(create=False):
    app = modal.App.lookup(APP, create_if_missing=True)
    try:
        return modal.Sandbox.from_name(APP, NAME)
    except modal.exception.NotFoundError:
        if not create:
            sys.exit("No sandbox; run: remote.py up")
        return modal.Sandbox.create(app=app, image=image, name=NAME, timeout=24 * 3600, cpu=8, memory=16384)


def run(sb, command, check=True):
    process = sb.exec("bash", "-lc", command, timeout=3 * 3600)
    for line in process.stdout:
        print(line, end="", flush=True)
    error = process.stderr.read()
    if error:
        print(error, end="", file=sys.stderr, flush=True)
    code = process.wait()
    if check and code:
        sys.exit(code)
    return code


def sync(sb, repo):
    root = os.path.join(CODE, repo)
    files = subprocess.run(
        ["git", "-C", root, "ls-files", "-co", "--exclude-standard", "-z"], check=True, capture_output=True
    ).stdout.split(b"\0")
    buffer = io.BytesIO()
    with tarfile.open(fileobj=buffer, mode="w:gz") as archive:
        for name in filter(None, files):
            path = os.path.join(root, name.decode())
            if os.path.isfile(path) and not os.path.islink(path):
                archive.add(path, arcname=name.decode())
        # Builds read git (the office template stamps `git describe`).
        if os.path.isdir(os.path.join(root, ".git")):
            archive.add(os.path.join(root, ".git"), arcname=".git")
    sb.filesystem.write_bytes(buffer.getvalue(), f"/tmp/{repo}.tgz")
    run(sb, f"mkdir -p /work/{repo} && tar xzf /tmp/{repo}.tgz -C /work/{repo}")
    print(f"synced {repo}: {len(buffer.getvalue()) // 1024} KiB")


def pull(sb, repo, path):
    run(sb, f"cd /work/{repo} && tar czf /tmp/pull.tgz {path}")
    data = sb.filesystem.read_bytes("/tmp/pull.tgz")
    target = os.path.join(CODE, repo)
    with tarfile.open(fileobj=io.BytesIO(data), mode="r:gz") as archive:
        archive.extractall(target, filter="data")
    print(f"pulled {repo}/{path}: {len(data) // 1024} KiB")


def main():
    action, *rest = sys.argv[1:]
    if action == "up":
        sb = sandbox(create=True)
        run(sb, "node -v && pnpm -v && cargo -V && wasm-pack -V")
    elif action == "down":
        sandbox().terminate()
    elif action == "sync":
        sb = sandbox()
        for repo in rest:
            sync(sb, repo)
    elif action == "run":
        sys.exit(run(sandbox(), f"cd /work/{rest[0]} && {rest[1]}", check=False))
    elif action == "pull":
        pull(sandbox(), rest[0], rest[1])
    else:
        sys.exit(__doc__)


main()
