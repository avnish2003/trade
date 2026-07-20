import express from "express";
import cors from "cors";
import { DatabaseSync } from "node:sqlite";
import path from "path";
import { fileURLToPath } from "url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const db = new DatabaseSync(path.join(__dirname, "journal.db"));

db.exec("PRAGMA journal_mode = WAL;");
db.exec(`
  CREATE TABLE IF NOT EXISTS kv (
    scope TEXT NOT NULL,
    key TEXT NOT NULL,
    value TEXT,
    updated_at TEXT NOT NULL,
    PRIMARY KEY (scope, key)
  );
`);

const getStmt = db.prepare("SELECT value FROM kv WHERE scope = ? AND key = ?");
const setStmt = db.prepare(`
  INSERT INTO kv (scope, key, value, updated_at) VALUES (?, ?, ?, ?)
  ON CONFLICT(scope, key) DO UPDATE SET value = excluded.value, updated_at = excluded.updated_at
`);

const app = express();
app.use(cors());
app.use(express.json({ limit: "25mb" }));

app.get("/api/kv/:key", (req, res) => {
  const scope = req.query.scope === "global" ? "global" : "user";
  const row = getStmt.get(scope, req.params.key);
  res.json({ value: row ? row.value : null });
});

app.put("/api/kv/:key", (req, res) => {
  const scope = req.body.scope === "global" ? "global" : "user";
  const value = req.body.value ?? "";
  setStmt.run(scope, req.params.key, value, new Date().toISOString());
  res.json({ ok: true });
});

app.get("/api/health", (req, res) => res.json({ ok: true }));

const PORT = 4000;
app.listen(PORT, () => console.log(`Trade journal API + SQLite listening on http://localhost:${PORT}`));
