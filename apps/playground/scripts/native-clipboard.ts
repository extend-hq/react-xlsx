import { execFile } from "node:child_process";
import { randomBytes } from "node:crypto";
import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import type { Plugin } from "vite";

const MAX_BODY_BYTES = 48 * 1024 * 1024;
const LOOPBACK = new Set(["127.0.0.1", "::1", "::ffff:127.0.0.1"]);
const LOCAL_HOSTS = new Set(["localhost", "127.0.0.1", "[::1]"]);

function run(executable: string, args: string[], input?: string) {
  return new Promise<string>((resolve, reject) => {
    const child = execFile(executable, args, { windowsHide: true, timeout: input ? 15_000 : 60_000, maxBuffer: input ? MAX_BODY_BYTES : 64 * 1024 }, (error, stdout, stderr) => {
      if (error) reject(new Error(stderr.trim() || error.message));
      else resolve(stdout);
    });
    child.stdin?.on("error", () => {});
    child.stdin?.end(input);
  });
}

export function nativeClipboard(): Plugin {
  let token = "";
  let directory: string | undefined;
  let executable: Promise<string> | undefined;
  let queue: Promise<unknown> = Promise.resolve();
  let pending = 0;
  let closed = false;
  const macHelper = () => executable ??= (async () => {
    directory = await mkdtemp(path.join(tmpdir(), "react-xlsx-clipboard-"));
    const binary = path.join(directory, "clipboard");
    await run("/usr/bin/xcrun", ["swiftc", path.resolve(__dirname, "native-clipboard.swift"), "-o", binary]);
    return binary;
  })();
  const invoke = async (body: string) => {
    if (process.platform === "darwin") return run(await macHelper(), [], body);
    if (process.platform === "win32") {
      const source = path.resolve(__dirname, "native-clipboard.cs").replaceAll("'", "''");
      const command = `$ErrorActionPreference = 'Stop'; Add-Type -Path '${source}' -ReferencedAssemblies System.dll,System.Windows.Forms.dll,System.Drawing.dll,System.Web.Extensions.dll; [XlsxNativeClipboard]::Run()`;
      const powershell = path.join(process.env.SystemRoot ?? "C:\\Windows", "System32", "WindowsPowerShell", "v1.0", "powershell.exe");
      return run(powershell, ["-NoLogo", "-NoProfile", "-NonInteractive", "-STA", "-Command", command], body);
    }
    throw new Error("Native clipboard integration is unavailable on this platform.");
  };
  return {
    name: "react-xlsx-native-clipboard",
    config(_config, environment) {
      token = environment.command === "serve" && !environment.isPreview && ["darwin", "win32"].includes(process.platform)
        ? randomBytes(32).toString("hex") : "";
      return { define: { __XLSX_NATIVE_CLIPBOARD_TOKEN__: JSON.stringify(token) } };
    },
    configureServer(server) {
      if (!token) return;
      server.httpServer?.once("close", () => {
        closed = true;
        void queue.finally(async () => {
          await executable?.catch(() => {});
          if (directory) await rm(directory, { recursive: true, force: true });
        });
      });
      server.middlewares.use("/__xlsx_native_clipboard", async (request, response) => {
        const fail = (status: number, message: string) => {
          response.writeHead(status, { "Content-Type": "application/json", "Cache-Control": "no-store" });
          response.end(JSON.stringify({ error: message }));
        };
        let origin: URL;
        try { origin = new URL(request.headers.origin ?? ""); }
        catch { fail(403, "A local same-origin request is required."); return; }
        if (!LOOPBACK.has(request.socket.remoteAddress ?? "") || !LOCAL_HOSTS.has(origin.hostname)
          || origin.host !== request.headers.host || !["http:", "https:"].includes(origin.protocol)
          || request.headers["x-xlsx-clipboard-token"] !== token) {
          fail(403, "A local same-origin request is required."); return;
        }
        if (request.method !== "POST" || request.headers["content-type"] !== "application/json") {
          fail(405, "Send clipboard data as JSON using POST."); return;
        }
        if (closed || pending >= 2) { fail(503, "The clipboard bridge is busy. Try copying again."); return; }
        if (Number(request.headers["content-length"]) > MAX_BODY_BYTES) { fail(413, "The drawing is too large to copy."); return; }
        pending++;
        try {
          const chunks: Buffer[] = [];
          let length = 0;
          for await (const chunk of request) {
            length += chunk.length;
            if (length > MAX_BODY_BYTES) { fail(413, "The drawing is too large to copy."); return; }
            chunks.push(chunk);
          }
          const body = Buffer.concat(chunks).toString("utf8");
          const payload = JSON.parse(body);
          if (payload.operation === "read") {
            const read = queue.then(() => invoke(body));
            queue = read.catch(() => {});
            const result = await read;
            response.writeHead(200, { "Content-Type": "application/json", "Cache-Control": "no-store" });
            response.end(result);
            return;
          }
          if (payload.operation !== "write") throw new Error("Invalid clipboard operation.");
          if (!["chart", "image"].includes(payload.kind)) throw new Error("Invalid drawing kind.");
          const decode = (key: string) => {
            const value = payload[key];
            if (typeof value !== "string" || !value.length || !/^[A-Za-z0-9+/]*={0,2}$/.test(value)) throw new Error("Invalid clipboard data.");
            return Buffer.from(value, "base64");
          };
          const drawing = decode("officeDrawing");
          const preview = decode("previewImage");
          const viewer = JSON.parse(decode("viewerDrawing").toString("utf8"));
          if (!drawing.subarray(0, 4).equals(Buffer.from([0x50, 0x4b, 3, 4])) || drawing.length > 32 * 1024 * 1024
            || !preview.subarray(0, 8).equals(Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]))
            || typeof viewer.shadowWorkbookBase64 !== "string" || viewer.floatingObjectType !== (payload.kind === "chart" ? 0 : 1)) {
            throw new Error("Invalid clipboard representations.");
          }
          const write = queue.then(() => invoke(body));
          queue = write.catch(() => {});
          await write;
          response.writeHead(204, { "Cache-Control": "no-store" });
          response.end();
        } catch (error) {
          server.config.logger.error(`Native clipboard: ${error instanceof Error ? error.message : String(error)}`);
          fail(500, "Could not access the native clipboard. Check the playground terminal for details.");
        } finally { pending--; }
      });
    },
  };
}
