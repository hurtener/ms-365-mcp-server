# Stateless HTTP Sidecar Contract

This fork supports a stateless Streamable HTTP mode intended for pooled HTTP sidecars.

## Summary

- Endpoint: `GET /mcp` and `POST /mcp`
- Required header: `Authorization: Bearer <Microsoft Graph access token>`
- Optional header: `x-microsoft-refresh-token: <Microsoft refresh token>`
- Session stickiness: not required
- Process reuse across users: supported
- `MS365_MCP_OAUTH_TOKEN`: not required for sidecar operation

## Request Authentication

Sidecar clients should authenticate every MCP request with a per-user Microsoft Graph bearer token.

Required header:

```http
Authorization: Bearer <access_token>
```

Optional header:

```http
x-microsoft-refresh-token: <refresh_token>
```

The refresh token header is only a convenience for the current request lifecycle. The server does not persist request
tokens across users or promote them into process-global auth state.

## Stateless Execution Model

- The HTTP transport runs in stateless mode.
- Each `/mcp` request creates an isolated MCP server instance and isolated request auth context.
- Request-scoped auth takes precedence over any other token source.
- In HTTP mode, requests without a bearer token are rejected and do not fall back to cached device-code auth.

This means a sidecar pool can safely reuse the same process for different users as long as the caller injects the
correct bearer token on every request.

## Recommended Sidecar Integration

Pooled HTTP consumers should treat Microsoft sidecars the same way they treat other bearer-auth MCP sidecars:

1. Run the server with HTTP mode enabled.
2. Send MCP requests to `/mcp`.
3. Inject the current user’s Graph access token in the `Authorization` header.
4. Optionally inject the refresh token if the sidecar should refresh access tokens mid-request.
5. Do not rely on startup environment tokens for sidecar routing.

Example:

```http
POST /mcp HTTP/1.1
Host: localhost:3000
Content-Type: application/json
Authorization: Bearer eyJ...
x-microsoft-refresh-token: 0.A...

{
  "jsonrpc": "2.0",
  "id": 1,
  "method": "tools/list",
  "params": {}
}
```

## Non-Goals

- Stdio mode continues to use the existing login/device-code and BYOT flows.
- `MS365_MCP_OAUTH_TOKEN` remains supported for non-sidecar consumers.
- OAuth discovery and dynamic registration remain available for generic MCP HTTP clients such as Open WebUI.

## Deployment Notes

- Sidecar pools can keep long-lived HTTP worker processes.
- No per-user process startup is required.
- No session affinity is required at the load balancer or pool layer.
- The sidecar is safe to reuse across users only when every request includes the correct bearer token.
