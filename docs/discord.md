# Discord integration

The project's [Discord server](https://discord.gg/haE8xpMET9) is wired to this repository so the
community sees what is happening without watching GitHub. Nothing here is required to build, test or release the package: every
Discord step is skipped when its webhook is not configured (forks, new clones).

## What posts where

| Channel          | Source                                       | Events                                                              |
| ---------------- | -------------------------------------------- | ------------------------------------------------------------------- |
| `#github-prs`    | repository webhook (Discord's GitHub format) | pull requests, reviews, review comments                             |
| `#github-issues` | repository webhook (Discord's GitHub format) | issues, issue and PR conversation comments, discussions and replies |
| `#releases`      | `announce-discord` job in `release.yml`      | one summary per release run, with the core, ui and mcp notes        |
| `#ci-alerts`     | `.github/workflows/discord-alerts.yml`       | CI, Release or Security failing on `main`                           |

Releases are announced by a job of our own rather than GitHub's release event because one run can
tag a dozen packages (the core, the UI, the MCP server and every viewer package); the job posts a
single message listing them all and says whether npm publishing succeeded. Mentions are disabled in
every payload, so commit subjects in the notes can never ping anyone.

## Setting it up (maintainers)

1. In Discord, open the channel's settings, Integrations, Webhooks, create one and copy its URL.
2. Run the matching command while the URL is still on the clipboard. The URL is read from the
   clipboard and handed to `gh`; it is never printed, passed as an argument or written to disk.

```sh
node scripts/discord-webhook.mjs prs        # repository webhook for #github-prs
node scripts/discord-webhook.mjs issues     # repository webhook for #github-issues
node scripts/discord-webhook.mjs releases   # Actions secret DISCORD_WEBHOOK_RELEASES
node scripts/discord-webhook.mjs ci         # Actions secret DISCORD_WEBHOOK_CI
```

Re-running `prs` or `issues` replaces that webhook instead of adding a second one. To rotate a
leaked URL, delete the webhook in Discord, create a new one and run the command again.

Preview a release message without sending it, from a checkout with a `release-plan.json`:

```sh
node scripts/announce-discord.mjs --plan release-plan.json --dry-run
```

## Moderation

Server rules live in `#rules` and point at the [Code of Conduct](../CODE_OF_CONDUCT.md). Security
reports go through the private advisory form in [SECURITY.md](../SECURITY.md), never Discord.
