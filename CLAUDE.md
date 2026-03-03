# Claude Code Memory - MCP Microsoft Office

## E2E Testing

### Run Tests

```bash
# Start the server (terminal 1)
npm run dev:web

# Run all tests (terminal 2)
node tests/run-all.cjs

# Run a single module
node tests/run-all.cjs --bucket mail --buckets-only

# Run multiple modules
node tests/run-all.cjs --bucket mail,calendar --buckets-only

# Run only workflow tests
node tests/run-all.cjs --workflows-only
```

### Test Structure

```
tests/
  lib/           Shared infrastructure (auth, http-client, reporter, harness)
  buckets/       One test file per module (9 files covering 78 tools)
  workflows/     Cross-module integration tests (5 files)
  run-all.cjs    Master test runner
  _archive/      Old test files (reference only)
```

### How Auth Works in Tests

Tests authenticate via ROPC (Resource Owner Password Credentials) — no manual token management. The test harness in `tests/lib/auth.cjs` handles:

1. ROPC call to Azure AD to get a Graph token
2. Exchange Graph token for MCP JWT via `POST /api/auth/graph-token-exchange`
3. Use MCP JWT as Bearer token for all API calls

Three test users are configured in `tests/lib/config.cjs`.

### Important Notes

- **tests/ is gitignored**: test files stay local, don't try to commit them
- **Server rate limits**: the in-memory rate limiter has a 15-minute window. Restart the server between rapid test runs.
- **Multi-day events**: calendarView API returns events that OVERLAP with the date range, not just events starting on that date

### API Parameter Reference

- `POST /calendar/events` body.contentType: lowercase `'text'` or `'html'`
- `POST /calendar/availability`: expects `{ users: [emails], timeSlots: [...] }`
- `POST /calendar/events/:id/accept|tentatively|decline`: body is `{ comment: string }`
- `POST /files/upload`: expects `{ name, content }`
- Files content/sharing endpoints: use `fileId` (not `id`) in request bodies
- `GET /v1/mail` returns a raw array, not `{ emails: [...] }`
- `GET /v1/mail/attachments` expects query param `id`, not `messageId`
- `POST /v1/mail/flag` expects `{ id, flag: true }`
- `POST /v1/mail/:id/reply` expects `{ body }`, not `{ comment }`
