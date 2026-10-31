const http = require("http");
const fs = require("fs");
const path = require("path");
const { URL } = require("url");

const PORT = Number(process.env.PORT) || 3000;
const PROJECT_ROOT = path.resolve(__dirname, "..");
const MIME_TYPES = {
  ".html": "text/html; charset=utf-8",
  ".css": "text/css; charset=utf-8",
  ".js": "application/javascript; charset=utf-8",
  ".json": "application/json; charset=utf-8",
  ".svg": "image/svg+xml",
  ".png": "image/png",
  ".jpg": "image/jpeg",
  ".ico": "image/x-icon",
};

function sendJson(res, status, payload) {
  res.writeHead(status, { "Content-Type": "application/json; charset=utf-8" });
  res.end(JSON.stringify(payload));
}

function serveFile(res, pathname) {
  const requestedPath = pathname === "/" ? "/frontend/index.html" : `/frontend${pathname}`;
  const filePath = path.normalize(path.join(PROJECT_ROOT, requestedPath));

  if (!filePath.startsWith(PROJECT_ROOT)) {
    return sendJson(res, 403, { error: "Forbidden" });
  }

  fs.stat(filePath, (error, stats) => {
    if (error || !stats.isFile()) {
      return sendJson(res, 404, { error: "File not found" });
    }

    const contentType = MIME_TYPES[path.extname(filePath).toLowerCase()] || "application/octet-stream";
    res.writeHead(200, {
      "Content-Type": contentType,
      "Cache-Control": "no-cache",
    });
    fs.createReadStream(filePath).pipe(res);
  });
}

const server = http.createServer((req, res) => {
  const requestUrl = new URL(req.url, `http://${req.headers.host || "localhost"}`);

  if (requestUrl.pathname === "/api/health") {
    return sendJson(res, 200, {
      status: "ok",
      service: "Little Stitches",
      timestamp: new Date().toISOString(),
    });
  }

  if (req.method !== "GET") {
    return sendJson(res, 405, { error: "Method not allowed" });
  }

  serveFile(res, requestUrl.pathname);
});

server.listen(PORT, () => {
  console.log(`Little Stitches berjalan di http://localhost:${PORT}`);
});
