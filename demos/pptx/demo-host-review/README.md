# Review hub (host-owned collaboration example)

A small "business system" in which several people review one presentation. The host owns
everything except the editor: sign-in, who may edit or only watch, the Yjs session and relay,
saved versions and the AI key. The PowerPoint editor (`pptx-vanilla-viewer`) is mounted with the
role and readiness the host grants. The recipe this implements is
[Review inside your application](../../../viewers/pptx/docs/guide/review-workflow.md).

```sh
bun install          # once, at the repository root
bun run start        # in this folder: http://localhost:4190 (PORT to change it)
bun run test         # server policy, tickets and an end-to-end relay test
```

Open the page in two browsers (or a normal and a private window) and sign in as different
people. Ana owns the review, Ben and Chen review, Dee watches read-only. As Ana, change someone's
access in the side panel: their editor switches role, or closes, without a reload.

The AI assistant appears for Ana and Ben (the accounts with an AI licence) once a model is
configured on the server:

```sh
REVIEW_AI_MODEL=gpt-4o-mini REVIEW_AI_API_KEY=sk-... bun run start
REVIEW_AI_MODEL=llama3.1 REVIEW_AI_BASE_URL=http://localhost:11434/v1 bun run start   # local, no key
```

The route needs `ooxml-ui`'s build (`bun run --cwd src/ui build` at the root) for its deck tools.

## Where each piece lives

| File                     | What it shows                                                                    |
| ------------------------ | -------------------------------------------------------------------------------- |
| `server/directory.ts`    | The host's accounts, permissions and versions, and the permission-to-role map    |
| `server/tickets.ts`      | Signed sign-in cookies and short-lived room tickets                              |
| `server/relay.ts`        | A y-websocket compatible relay that admits only ticket holders                   |
| `server/relay-policy.ts` | Drops document writes from read-only sockets; pins awareness ids to connections  |
| `server/app.ts`          | The API, the seed lease, owner-only versions, access changes, the AI gate        |
| `server/ai-route.ts`     | The model route with schema-only deck tools; the key stays on the server         |
| `src/review-session.ts`  | The host-owned `externalSession`: readiness, ticket refresh, access-change close |
| `src/review-page.ts`     | Mounting, remounting on an access change, owner write-back, the side panel       |

## Limits

Everything is in memory (restart and it is gone) and sign-in is an account picker. The owner's
editor writes a version after changes settle, and also once after it mounts; nobody saves while
the owner is offline. Awareness payloads (name, colour) are not checked against the ticket. The
AI route is tested against a mock OpenAI-compatible server, not a real model.
