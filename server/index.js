import express from "express";
import cors from "cors";
import { neon } from "@neondatabase/serverless";

process.loadEnvFile(new URL("../.env", import.meta.url));

const sql = neon(process.env.DATABASE_URL);

await sql`
  CREATE TABLE IF NOT EXISTS kv (
    scope TEXT NOT NULL,
    key TEXT NOT NULL,
    value TEXT,
    updated_at TIMESTAMPTZ NOT NULL,
    PRIMARY KEY (scope, key)
  )
`;

const app = express();
app.use(cors());
app.use(express.json({ limit: "25mb" }));

app.get("/api/kv/:key", async (req, res) => {
  const scope = req.query.scope === "global" ? "global" : "user";
  const rows = await sql`SELECT value FROM kv WHERE scope = ${scope} AND key = ${req.params.key}`;
  res.json({ value: rows[0]?.value ?? null });
});

app.put("/api/kv/:key", async (req, res) => {
  const scope = req.body.scope === "global" ? "global" : "user";
  const value = req.body.value ?? "";
  await sql`
    INSERT INTO kv (scope, key, value, updated_at) VALUES (${scope}, ${req.params.key}, ${value}, now())
    ON CONFLICT (scope, key) DO UPDATE SET value = excluded.value, updated_at = excluded.updated_at
  `;
  res.json({ ok: true });
});

app.get("/api/health", (req, res) => res.json({ ok: true }));

const PORT = 4000;
app.listen(PORT, () => console.log(`Trade journal API + Neon Postgres listening on http://localhost:${PORT}`));
