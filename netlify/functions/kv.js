import { neon } from "@neondatabase/serverless";

const sql = neon(process.env.DATABASE_URL);

async function ensureTable() {
  await sql`
    CREATE TABLE IF NOT EXISTS kv (
      scope TEXT NOT NULL,
      key TEXT NOT NULL,
      value TEXT,
      updated_at TIMESTAMPTZ NOT NULL,
      PRIMARY KEY (scope, key)
    )
  `;
}

export default async (req, context) => {
  await ensureTable();
  const key = context.params.key;
  const url = new URL(req.url);

  if (req.method === "GET") {
    const scope = url.searchParams.get("scope") === "global" ? "global" : "user";
    const rows = await sql`SELECT value FROM kv WHERE scope = ${scope} AND key = ${key}`;
    return new Response(JSON.stringify({ value: rows[0]?.value ?? null }), {
      headers: { "content-type": "application/json" },
    });
  }

  if (req.method === "PUT") {
    const body = await req.json();
    const scope = body.scope === "global" ? "global" : "user";
    const value = body.value ?? "";
    await sql`
      INSERT INTO kv (scope, key, value, updated_at) VALUES (${scope}, ${key}, ${value}, now())
      ON CONFLICT (scope, key) DO UPDATE SET value = excluded.value, updated_at = excluded.updated_at
    `;
    return new Response(JSON.stringify({ ok: true }), {
      headers: { "content-type": "application/json" },
    });
  }

  return new Response("Method not allowed", { status: 405 });
};

export const config = { path: "/api/kv/:key" };
