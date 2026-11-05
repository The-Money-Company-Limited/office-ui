// The build box: a Worker whose Builder Durable Object owns one container. GitHub Actions deploys it
// with `cf deploy` (.github/workflows/office-builds.yml), building the image there with Docker.
import { bindings, defineConfig, defineContainer, defineWorker, exports } from "cf/config";

// The name the box's container application has had since it was first deployed.
const box = defineContainer({
  name: "office-builds-builder",
  schedulingPolicy: "durable-object",
  // src/index.ts starts `base` with the standard-4 instance type.
  images: { base: { dockerfile: "./Dockerfile" } },
});

const worker = defineWorker({
  name: "office-builds",
  compatibilityDate: "2026-11-05",
  entrypoint: "src/index.ts",
  observability: { enabled: true },
  env: {
    BUILDER: bindings.durableObject({ worker: "office-builds", exportName: "Builder" }),
  },
  exports: {
    Builder: exports.durableObject({ storage: "sqlite", container: box }),
  },
});

export default defineConfig({ worker, containers: [box] });
