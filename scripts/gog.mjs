import { mkdirSync } from "node:fs";
import { join } from "node:path";
import { spawnSync } from "node:child_process";

function run(cmd, args) {
  const res = spawnSync(cmd, args, { stdio: "inherit" });
  if (res.error) throw res.error;
  if (res.status !== 0) {
    process.exit(typeof res.status === "number" ? res.status : 1);
  }
  return res.status;
}

function runCapture(cmd, args) {
  const res = spawnSync(cmd, args, { stdio: ["ignore", "pipe", "ignore"] });
  if (res.error || res.status !== 0) return "";
  return String(res.stdout || "").trim();
}

const repoRoot = process.cwd();
const binDir = join(repoRoot, "bin");
mkdirSync(binDir, { recursive: true });

const exe = process.platform === "win32" ? "gog.exe" : "gog";
const binPath = join(binDir, exe);

const version = runCapture("git", ["describe", "--tags", "--always", "--dirty"]) || "dev";
const commit = runCapture("git", ["rev-parse", "--short=12", "HEAD"]) || "";
const date = new Date().toISOString().replace(/\.\d{3}Z$/, "Z");
const ldflags = [
  `-X github.com/steipete/gogcli/internal/cmd.version=${version}`,
  `-X github.com/steipete/gogcli/internal/cmd.commit=${commit}`,
  `-X github.com/steipete/gogcli/internal/cmd.date=${date}`,
].join(" ");

run("go", ["build", "-ldflags", ldflags, "-o", binPath, "./cmd/gog"]);

const final = spawnSync(binPath, process.argv.slice(2), { stdio: "inherit" });
if (final.error) throw final.error;
process.exit(typeof final.status === "number" ? final.status : 1);
