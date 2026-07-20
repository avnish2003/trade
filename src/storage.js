// Shim for the window.storage.get/set interface the journal component expects,
// backed by the real SQLite database via the local Express API instead of the browser.
async function get(key, isGlobal = false) {
  const res = await fetch(`/api/kv/${encodeURIComponent(key)}?scope=${isGlobal ? "global" : "user"}`);
  if (!res.ok) return { value: null };
  return res.json();
}

async function set(key, value, isGlobal = false) {
  await fetch(`/api/kv/${encodeURIComponent(key)}`, {
    method: "PUT",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ value, scope: isGlobal ? "global" : "user" }),
  });
}

window.storage = { get, set };
