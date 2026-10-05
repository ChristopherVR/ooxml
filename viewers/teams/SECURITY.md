# Security Policy

## Supported versions

Each published package in this monorepo is versioned and released independently
(see [CONTRIBUTING.md](./CONTRIBUTING.md#commit-conventions) for the release process). Security
fixes target the **latest release line of each package**; we do
not maintain long-lived security-fix branches for older majors.

| Package            | npm name                   |
| ------------------ | -------------------------- |
| `packages/react`   | `openteams-react-viewer`   |
| `packages/vue`     | `openteams-vue-viewer`     |
| `packages/angular` | `openteams-angular-viewer` |
| `packages/svelte`  | `openteams-svelte-viewer`  |
| `packages/solid`   | `openteams-solid-viewer`   |
| `packages/vanilla` | `openteams-vanilla-viewer` |
| `server`           | `openteams-server`         |

The OpenTeams logic lives in [`ooxml-core`](https://github.com/ChristopherVR/ooxml) (`ooxml-core/teams`) and the
`<teams-app>` element in `ooxml-ui/teams`; both are consumed from there. A vulnerability in either is fixed in
the `ooxml` repository and reaches this one through a new release, so it is reported there if
that is where the affected code lives (see its [SECURITY.md](https://github.com/ChristopherVR/ooxml/blob/main/SECURITY.md)).

Always upgrade to the latest version of the package(s) you depend on before
reporting a suspected vulnerability, in case it has already been fixed.

## Reporting a vulnerability

**Please do not open a public GitHub issue for security vulnerabilities.**

Report vulnerabilities privately using
[GitHub's private vulnerability reporting](https://github.com/ChristopherVR/teams-viewer/security/advisories/new)
(the "Report a vulnerability" button under this repository's **Security** tab).
This opens a private advisory visible only to the maintainers until a fix is
ready, and lets you attach a proof of concept without exposing it publicly.

If you're unable to use GitHub's private reporting for any reason, open a
regular issue asking a maintainer to reach out and provide the details out of
band; do not include exploit details in the issue itself.

Please include as much of the following as you can:

- The affected package(s) and version(s).
- A description of the vulnerability and its potential impact.
- Steps to reproduce, ideally a minimal repro or a code snippet - this project handles
  untrusted network input (the sync protocol, signaling messages, chat content, attachments) and the reference server, so authorisation bypass, message spoofing, hangs, memory exhaustion,
  and XSS via rendered chat or channel content are all in scope, not just "classic" injection bugs.
- Whether the issue is reachable by an unauthenticated peer, or needs an existing account or room membership.

## Response expectations

This is a side project maintained outside of full-time work, not a funded or
staffed effort, so there's no fixed SLA on response times - reports are
triaged and fixed when maintainer time allows rather than on a guaranteed
schedule. That said, we do take reports seriously and will get to them.

Once a fix is ready, we'll coordinate a disclosure timeline with the reporter
before any public advisory or changelog entry is published. We credit
reporters in the advisory unless you ask to remain anonymous.

## Scope

In scope:

- All packages under `packages/` and their published npm artifacts.
- The demo apps under `demo/` or `demos/` only insofar as a bug there reveals a
  vulnerability in one of the published packages they consume (the demos
  themselves are not published or deployed as a security-relevant surface).

Out of scope:

- Vulnerabilities that require the operator and every participant to already fully
  trust each other, with no untrusted peer or content boundary crossed.
- Findings from automated scanners without a demonstrated, concrete impact
  (e.g. a generic dependency CVE with no reachable code path in this project).
  Dependency vulnerabilities are tracked through the automated `ooxml` adoption workflow and
  normal dependency updates; feel free to open a normal issue for those.

## Automated scanning

Code-scanning alerts, when available, can be reviewed on GitHub; the checked-in workflows do not define a CodeQL job.
Findings are triaged and fixed as part of normal development; you don't need
to separately report something that's already visible in the repository's
public [code scanning alerts](https://github.com/ChristopherVR/teams-viewer/security/code-scanning),
though private vulnerability reporting is still preferred for anything with a
working exploit.
