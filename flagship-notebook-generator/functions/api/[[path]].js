export async function onRequest(context) {
  const backendBase = (context.env.BACKEND_BASE_URL || "").replace(/\/$/, "");
  if (!backendBase) {
    return jsonError("Missing BACKEND_BASE_URL environment variable.", 500);
  }

  const incoming = new URL(context.request.url);
  const suffix = incoming.pathname.replace(/^\/api/, "");
  const target = new URL(`${backendBase}/api${suffix}${incoming.search}`);
  const headers = new Headers(context.request.headers);
  headers.delete("host");

  const response = await fetch(target.toString(), {
    method: context.request.method,
    headers,
    body: ["GET", "HEAD"].includes(context.request.method) ? undefined : context.request.body,
    redirect: "follow",
  });

  return new Response(response.body, {
    status: response.status,
    statusText: response.statusText,
    headers: response.headers,
  });
}

function jsonError(message, status) {
  return new Response(JSON.stringify({ error: message }), {
    status,
    headers: {
      "content-type": "application/json; charset=utf-8",
    },
  });
}
