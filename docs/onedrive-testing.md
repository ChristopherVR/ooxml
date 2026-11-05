# Testing OneDrive integration

The shared connection is exercised by automated, test-local Graph transports. The suite tests
exercise the same connection contracts across Word, Excel, PowerPoint and Visio. Test fixtures
do not contact Microsoft, publish sharing links or stand in for live account acceptance. There
is no shipped OneDrive simulation service, public mock export or simulation demo.

## Automated checks

From the repository root with the CI version of Bun:

```sh
bun install --frozen-lockfile
bun run test -- embedding/onedrive
bun run --cwd src/core typecheck:strict
bun run build:suite
bun run preview:suite -- 8146
# In a second terminal:
bun run test:suite:onedrive
```

The targeted suite browser check defaults to `http://127.0.0.1:8146`; set `SUITE_URL` to use a
different running preview. Its network fixtures are test-only. Root `bun run verify` selects
checks from the changed paths; use `--browser` when verifying changes to editor wiring.

The meaningful cases are file identities and native MIME/type admission for all four formats;
token boundaries on Graph versus preauthenticated transfer requests; serialized saves; revision-only
conflicts; true concurrent content changes; cloud-copy recovery leaving the original unchanged;
native save events; pending-cloud persistence; read-only restrictions; auth errors; cancellation;
and sharing roles/scopes. Failures must preserve local bytes and must not report a successful cloud save.

## What can be tested without our own app registration?

All transport and host-wiring regression tests can run without Microsoft credentials. They do
not validate OAuth, tenant policy, real service permissions, deployed browser CORS, sensitivity
labels or interoperability with Microsoft Office.

For manual API exploration, Microsoft Graph Explorer uses Microsoft's existing application
registration. The [quick start](onedrive-quickstart.md#test-without-your-own-registration) shows
how to use its token to try a personal OneDrive file during development. Sign in with a test account and consent to permitted scopes. Anonymous sample data
does not represent your OneDrive. Follow [Microsoft's Graph Explorer guide](https://learn.microsoft.com/en-us/graph/graph-explorer/graph-explorer-overview).

Start with a disposable native Office file. Resolve its sharing link through
`GET /shares/{encodedLink}/driveItem`, inspect the drive/item/revision metadata, and download
its bytes. Test conditional upload and copy creation only against disposable files. Use Graph
Explorer's role and scope controls when checking link creation, then remove test permissions.

Graph Explorer checks its own application registration, not our application's consent, redirect
URIs, host token provider or deployed-origin behavior. Do not copy its client ID into production
or make the application depend on borrowing another tool's access token. An already authorized
organization host can supply its existing token provider, subject to its permissions and policy.

### Test a personal OneDrive account

The production host can support a personal Microsoft account with a registration configured for
personal users and a Graph delegated `Files.ReadWrite` token. Use `consumers` or the appropriate
`common` authority and your own registered redirect. A Microsoft account-management login URL
does not itself provide Graph file authorization. See the [personal account setup](onedrive-integration.md#personal-microsoft-accounts).

Use a disposable DOCX, XLSX, PPTX or VSDX in that account, resolve its link, edit it in the matching
suite editor, save, then reopen from OneDrive and in its native Office application. Check both
same-item saving and conflict-copy recovery. When trying Graph Explorer first, use the account
types and consent flow available in that tool; a successful API smoke test still does not validate
the suite's own sign-in setup. No personal login hint, session state or real authorization URL
belongs in repository documentation or test fixtures.

### Run the real-service smoke check

`scripts/onedrive-live-check.mjs` uses the same shared core client against an actual authorized
Graph token. It does not implement sign-in or bypass app consent. Obtain a short-lived token
from an authorized local developer host, or use Graph Explorer as part of a manual API test.
The token must target Microsoft Graph with the required file scopes; a Microsoft account-settings
session or token for another resource does not suffice.

Build core once, then run the check with process-local environment variables. Do not send the
token in chat, put it in source, or paste it as a command argument. For PowerShell 7:

```powershell
bun run build
node scripts/onedrive-live-check.mjs --help
$env:OOXML_ONEDRIVE_ACCESS_TOKEN = Read-Host 'Graph access token' -MaskInput
$env:OOXML_ONEDRIVE_SHARE_URL = Read-Host 'OneDrive sharing link' -MaskInput
try {
	node scripts/onedrive-live-check.mjs
} finally {
	Remove-Item Env:\OOXML_ONEDRIVE_ACCESS_TOKEN -ErrorAction SilentlyContinue
	Remove-Item Env:\OOXML_ONEDRIVE_SHARE_URL -ErrorAction SilentlyContinue
}
```

The default check resolves, downloads and parses the native document without writing it.
For a deliberate write test, run `node scripts/onedrive-live-check.mjs --write-copy` while those
environment variables are set. It creates a separate sibling copy, then downloads and validates
that copy; the original file is not replaced. The copy remains on the service so it can be opened
in the native Office application and removed when the test is complete. The script does not
print access tokens, sharing URLs or temporary transfer URLs.

### Measure download speed

`scripts/onedrive-speed-check.mjs` times the same download in layers: the Graph metadata call, the
download host (first byte, total, throughput, chunks), `client.open` / `openItem` with
`onProgress`, `createOneDriveFileProvider(...).open`, optionally Graph `/content` (302 redirect)
against metadata plus download URL, and `list()`, `recent()` and an optional folder. It prints
per-run values, median/min/max and a short reading, for example the library overhead over a raw
`fetch` of a fresh download URL. Tokens, links, paths and download URLs are never printed. The
token is static, so "token acquisition" is a function call, not a sign-in.

```powershell
bun run build
node scripts/onedrive-speed-check.mjs --help
$env:OOXML_ONEDRIVE_ACCESS_TOKEN = Read-Host 'Graph access token' -MaskInput
$env:OOXML_ONEDRIVE_PATH = Read-Host 'Path in your drive, e.g. Documents/deck.pptx' -MaskInput
try {
	node scripts/onedrive-speed-check.mjs --runs=5
} finally {
	Remove-Item Env:\OOXML_ONEDRIVE_ACCESS_TOKEN -ErrorAction SilentlyContinue
	Remove-Item Env:\OOXML_ONEDRIVE_PATH -ErrorAction SilentlyContinue
}
```

`OOXML_ONEDRIVE_SHARE_URL` (a sharing link) works instead of or together with the path, and
`OOXML_ONEDRIVE_FOLDER` adds a folder listing. Add `--json` for machine-readable output,
`--warmup` to discard a cold first run, `--no-content` to skip the `/content` comparison.

`bun run test:onedrive:speed` runs `--mock`: a local server imitates Graph and the download host
(`--size=MB`, `--latency=ms`, `--bandwidth=MBs`), needs no Microsoft access and gives a regression
baseline for the library's own overhead. It says nothing about Microsoft's network. Live numbers
depend on your connection, region, file size and time of day; compare medians over several runs,
and treat a library overhead within about 10% (or a few tens of milliseconds) as noise.

This smoke check is real service evidence for the supplied account, token and document. It still
does not validate the suite's browser sign-in integration, redirect URI, deployed CORS behavior,
every sharing policy or simultaneous Microsoft Office coauthoring.

## Real-account acceptance

Record account type, browser, authorized application, scopes, disposable document, observed
result and native Office reopen result. Keep tokens and capability URLs out of test reports.
Every row below remains a live acceptance requirement until run against the configured service.

| Scenario                                           | Required observation                                                                                               |
| -------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------ |
| DOCX, XLSX, PPTX and VSDX                          | Correct editor opens; native edit saves to the same drive/item; corresponding Microsoft app reopens output.        |
| Personal and work/school OneDrive                  | Supported account flow, consent and write permissions work for the actual host registration.                       |
| SharePoint-backed links                            | Correct item resolves under tenant and cross-tenant access policy.                                                 |
| Unknown/read-only access                           | Viewer restrictions apply; cloud Save is unavailable; original is not modified.                                    |
| Same content, newer eTag                           | Session refreshes baseline identity and retries once without inventing a binary merge.                             |
| Independently changed supported content            | Local and remote changes combine into one saved file and the editor adopts the result.                             |
| Overlapping or unsupported changes                 | Both versions retained; explicit cloud-copy/latest/local recovery remains available.                               |
| Cloud-copy recovery                                | Separate item created; original remote version unchanged; continued edits target the copy.                         |
| Revoked access, expired token or wrong account     | Honest failure, recoverable local bytes and pending state retained.                                                |
| Interrupted or canceled upload                     | No false saved state; latest revision inspected before any deliberate retry.                                       |
| Remote change after session creation               | Characterize service behavior; do not claim a lock through final commit.                                           |
| View/edit, organization/users/anonymous sharing    | Explicit choices honored or policy rejection shown; no anonymous fallback.                                         |
| OpenTeams link versus attachment copy              | Original OneDrive item retains its identity; copied attachment is presented as independent.                        |
| App reload and profile switch                      | Safe cloud identity/baseline/pending state retained appropriately; tokens and capability URLs absent from storage. |
| Representative charts, fonts, media and embeddings | Review native export diagnostics and compare in the relevant Microsoft application.                                |

These checks do not claim Microsoft coauthoring, merging of every Office feature, durable server locks or
universally lossless Office round trips. See [the integration contract](onedrive-integration.md).
