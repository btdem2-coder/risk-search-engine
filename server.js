require("dotenv").config();

const http = require("http");
const fs = require("fs");
const path = require("path");
const ExcelJS = require("exceljs");
const { askRiskSearch, getFilterOptions } = require("./ask");

const PORT = Number(process.env.PORT) || 3000;
const indexHtmlPath = path.join(__dirname, "public", "index.html");
const defaultLogoPath = path.join(__dirname, "public", "logo.png");
const APP_USER = process.env.APP_USER;
const APP_PASS = process.env.APP_PASS;
const APP_USERS = process.env.APP_USERS;
const LOGO_PATH = process.env.LOGO_PATH || defaultLogoPath;

function sendJson(res, statusCode, payload) {
  res.writeHead(statusCode, { "Content-Type": "application/json" });
  res.end(JSON.stringify(payload));
}

function parseBasicAuthHeader(headerValue) {
  if (!headerValue || !headerValue.startsWith("Basic ")) return null;
  const encoded = headerValue.slice(6).trim();
  if (!encoded) return null;

  try {
    const decoded = Buffer.from(encoded, "base64").toString("utf8");
    const separatorIndex = decoded.indexOf(":");
    if (separatorIndex < 0) return null;
    return {
      username: decoded.slice(0, separatorIndex),
      password: decoded.slice(separatorIndex + 1),
    };
  } catch {
    return null;
  }
}

function parseAppUsersList(rawValue) {
  if (!rawValue) return [];

  return String(rawValue)
    .split(",")
    .map((item) => item.trim())
    .filter(Boolean)
    .map((pair) => {
      const separatorIndex = pair.indexOf(":");
      if (separatorIndex <= 0) return null;
      const username = pair.slice(0, separatorIndex).trim();
      const password = pair.slice(separatorIndex + 1).trim();
      if (!username || !password) return null;
      return { username, password };
    })
    .filter(Boolean);
}

const configuredUsers = parseAppUsersList(APP_USERS);
if (APP_USER && APP_PASS) {
  configuredUsers.push({ username: APP_USER, password: APP_PASS });
}

function isAuthenticated(req) {
  // Auth is optional for local/dev convenience.
  if (!configuredUsers.length) return true;

  const credentials = parseBasicAuthHeader(req.headers.authorization);
  if (!credentials) return false;
  return configuredUsers.some(
    (user) =>
      user.username === credentials.username && user.password === credentials.password
  );
}

function sendUnauthorized(res) {
  res.writeHead(401, {
    "Content-Type": "application/json",
    "WWW-Authenticate": 'Basic realm="Aldar International Risk Search"',
  });
  res.end(JSON.stringify({ error: "Unauthorized." }));
}

function serveIndexHtml(res) {
  fs.readFile(indexHtmlPath, "utf8", (error, content) => {
    if (error) {
      sendJson(res, 500, { error: "Unable to load UI file." });
      return;
    }

    res.writeHead(200, { "Content-Type": "text/html; charset=utf-8" });
    res.end(content);
  });
}

function serveLogo(res) {
  fs.readFile(LOGO_PATH, (error, content) => {
    if (error) {
      sendJson(res, 404, {
        error: `Logo file not found at ${LOGO_PATH}. Add public/logo.png or set LOGO_PATH.`,
      });
      return;
    }

    res.writeHead(200, { "Content-Type": "image/png" });
    res.end(content);
  });
}

function parseJsonBody(req) {
  return new Promise((resolve, reject) => {
    let body = "";

    req.on("data", (chunk) => {
      body += chunk;
      if (body.length > 1_000_000) {
        reject(new Error("Payload too large."));
      }
    });

    req.on("end", () => {
      try {
        resolve(body ? JSON.parse(body) : {});
      } catch {
        reject(new Error("Invalid JSON payload."));
      }
    });

    req.on("error", () => reject(new Error("Failed reading request body.")));
  });
}

function stringifyValue(value) {
  if (Array.isArray(value)) {
    return value.join(", ");
  }
  if (value === null || value === undefined || value === "") {
    return "";
  }
  return String(value);
}

async function buildExportWorkbook(payload = {}) {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = "risk-search-console";
  workbook.created = new Date();

  const metadataSheet = workbook.addWorksheet("Search Context");
  metadataSheet.columns = [
    { header: "Field", key: "field", width: 30 },
    { header: "Value", key: "value", width: 80 },
  ];

  const mode = payload.mode === "section2" ? "Section 2 - Filtration Results" : "Section 1 - Prompt Search";
  const question = payload.question || "";
  const selectedFilters = payload.selectedFilters || {};
  const appliedFilters = payload.appliedFilters || {};

  metadataSheet.addRow({ field: "Mode", value: mode });
  if (payload.mode === "section1") {
    metadataSheet.addRow({ field: "Question", value: question });
    metadataSheet.addRow({
      field: "Selected Entity (Section 1)",
      value: stringifyValue(selectedFilters.entity),
    });
    metadataSheet.addRow({
      field: "Selected Industry (Section 1)",
      value: stringifyValue(selectedFilters.industry),
    });
    metadataSheet.addRow({
      field: "Applied Entity Filters",
      value: stringifyValue(appliedFilters.entity),
    });
    metadataSheet.addRow({
      field: "Applied Industry Filters",
      value: stringifyValue(appliedFilters.industry),
    });
  } else {
    metadataSheet.addRow({
      field: "Chosen Entities (Section 2)",
      value: stringifyValue(selectedFilters.entity),
    });
    metadataSheet.addRow({
      field: "Chosen Industries (Section 2)",
      value: stringifyValue(selectedFilters.industry),
    });
  }
  metadataSheet.addRow({ field: "Exported At", value: new Date().toISOString() });

  const resultsSheet = workbook.addWorksheet("Results");
  resultsSheet.columns = [
    { header: "#", key: "index", width: 8 },
    { header: "ID", key: "id", width: 28 },
    { header: "Risk Subject", key: "riskSubject", width: 48 },
    { header: "Score", key: "score", width: 14 },
    { header: "Industry", key: "industry", width: 30 },
    { header: "Entity", key: "entity", width: 40 },
    { header: "Risk Description", key: "riskDescription", width: 80 },
  ];

  const results = Array.isArray(payload.results) ? payload.results : [];
  results.forEach((item, idx) => {
    resultsSheet.addRow({
      index: idx + 1,
      id: item.id || "",
      riskSubject: item.riskSubject || "",
      score: item.score ?? "",
      industry: item.industry || "",
      entity: item.entity || "",
      riskDescription: item.riskDescription || "",
    });
  });

  return workbook.xlsx.writeBuffer();
}

const server = http.createServer(async (req, res) => {
  const { method, url } = req;

  if (method === "GET" && url === "/health") {
    sendJson(res, 200, { ok: true });
    return;
  }

  if (!isAuthenticated(req)) {
    sendUnauthorized(res);
    return;
  }

  if (method === "GET" && url === "/") {
    serveIndexHtml(res);
    return;
  }

  if (method === "GET" && url === "/logo.png") {
    serveLogo(res);
    return;
  }

  if (method === "GET" && url === "/api/options") {
    sendJson(res, 200, getFilterOptions());
    return;
  }

  if (method === "POST" && url === "/api/ask") {
    try {
      const body = await parseJsonBody(req);
      const question = body.question;
      const filters = body.filters || {};
      const response = await askRiskSearch(question, filters);
      sendJson(res, 200, response);
    } catch (error) {
      sendJson(res, 400, {
        error: error.response?.data || error.message || "Request failed.",
      });
    }
    return;
  }

  if (method === "POST" && url === "/api/export") {
    try {
      const body = await parseJsonBody(req);
      const buffer = await buildExportWorkbook(body);
      const filename = `risk-search-export-${Date.now()}.xlsx`;
      res.writeHead(200, {
        "Content-Type":
          "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "Content-Disposition": `attachment; filename="${filename}"`,
      });
      res.end(Buffer.from(buffer));
    } catch (error) {
      sendJson(res, 400, {
        error: error.response?.data || error.message || "Export failed.",
      });
    }
    return;
  }

  sendJson(res, 404, { error: "Not found." });
});

server.listen(PORT, () => {
  console.log(`UI server running on http://localhost:${PORT}`);
});
