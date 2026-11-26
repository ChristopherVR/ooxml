# Office add-ins and an Office.js-compatible runtime

Word, Excel, and PowerPoint will run compatible task-pane add-ins against the live editor. The
suite will use documented Microsoft interfaces for discovery and administration where available.
Hosts can supply their own catalogs, packages, and services. A Marketplace listing, a manifest,
and a package that can run in this host are three distinct deliverables.

## What exists today

- [`src/core/ribbon/add-in.ts`](../src/core/ribbon/add-in.ts) defines `RibbonAddInTab`, its
  groups and commands, and `office-ribbon-add-in`. It filters empty, duplicate, and built-in tab
  IDs. Commands can run trusted host callbacks.
- [`src/ui/src/ribbon/add-in-settings.ts`](../src/ui/src/ribbon/add-in-settings.ts) exposes
  `setOfficeAddInCatalog`. The host supplies IDs, labels, and install and optional remove
  callbacks. Admin Managed lists the host catalog, and My Add-ins discovers and activates
  mounted ribbon tabs. Hosts can use `setOfficeAddInStorefront` to enable the Store tab
  with a Marketplace link or a host-supplied list of offers.
  [`add-in-tabs.ts`](../src/ui/src/ribbon/add-in-tabs.ts) renders host tabs in Word and Excel.
  PowerPoint uses [`ribbon-add-ins.ts`](../src/ui/src/pptx/render/ribbon-add-ins.ts) and
  [`ribbon-add-in.ts`](../src/ui/src/pptx/web-components/ribbon-add-in.ts).
- [`apps/office-suite/src/editors.js`](../apps/office-suite/src/editors.js) mounts the editors.
  [`session.js`](../apps/office-suite/src/session.js) owns document sessions, revisions, save, and
  preview state. [`ai-operations.js`](../apps/office-suite/src/ai-operations.js) lazy-loads core
  automation for the assistant. Third-party task panes are open work.
- [`src/core/automation/index.ts`](../src/core/automation/index.ts) exposes byte-oriented Word
  and Excel operations and shared tool definitions. [`live.ts`](../src/core/automation/live.ts)
  serializes tools per registered editor and checks for changes before loading edited bytes.
  PowerPoint defines its tools in [`src/core/pptx/automation/catalogue/`](../src/core/pptx/automation/catalogue/).
  Office.js proxy objects, selection lifetime, and batch contexts are open work.

The checked catalog item in the [parity checklist](parity-checklist.md) covers trusted host tabs
and host-enabled Marketplace browsing. The checked Forms item covers response links in Excel and
PowerPoint and on-demand loading during a slide show. [`microsoft-forms.ts`](../src/ui/src/form/microsoft-forms.ts)
validates those links. Native embedding, signed-in management, and live workbook response sync
remain open. A separate BYO Forms provider can receive responses through a tenant-configured
Power Automate flow using Microsoft's documented Forms connector. That flow can consume new
responses and get their details. Native management and workbook sync still need their own product
and service contracts. [Forms connector](https://learn.microsoft.com/en-us/connectors/microsoftforms/)
· [PowerPoint embedding](https://support.microsoft.com/en-gb/forms/insert-a-form-or-quiz-into-powerpoint)
· [Excel live connections](https://support.microsoft.com/en-us/forms/can-i-get-real-time-updates-in-my-excel-workbook).

- Local install accepts task-pane XML, unified JSON and bounded unified ZIP packages in
  Word, Excel and PowerPoint. Upload is available from every tab of the shared dialog and
  previews declared publisher, version, origins and permissions before consent. My Add-ins
  lists and removes the in-memory installs. An invalid replacement leaves the installed version
  available in the list. Execution, persistent installs and remote storefront providers remain open.

## Parse two public manifest formats

Add a DOM-free manifest reader under `src/core/automation/add-ins/`. Both readers produce the
same typed descriptor with identity, version, publisher, localized metadata, Office hosts,
requirements, requested permissions, and resolved task-pane and command entry points. Keep
declared origins, icon references, and authentication metadata separate from granted policy.

- **Add-in-only XML.** Read namespaced `OfficeApp` with `xsi:type="TaskPaneApp"`. Resolve
  `Id`, `Version`, `ProviderName`, localized names, descriptions and URLs, and the `Document`,
  `Workbook`, and `Presentation` hosts. Read `DefaultSettings.SourceLocation`, `Requirements`,
  `Permissions`, and `AppDomains`. Resolve supported `VersionOverrides` 1.0 task-pane commands
  and their resource IDs. Apply base requirements first. Fall back to the base pane when the
  override requirements fail. [XML manifest overview](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/xml-manifest-overview)
  · [XML manifest reference](https://learn.microsoft.com/en-us/javascript/api/manifest)
  · [Requirements element](https://learn.microsoft.com/en-us/javascript/api/manifest/requirements).
- **Unified Microsoft 365 JSON.** Read versioned `manifest.json` and package-local icons and
  locale data. Resolve applicable Office `extensions`, `requirements.scopes` and `capabilities`,
  `runtimes`, actions, and `ribbons`. Read `validDomains`,
  `authorization.permissions.resourceSpecific`, and optional web application identity.
  Evaluate requirements at extension and feature scope. Report other app extensions as inactive.
  [Unified manifest overview](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/unified-manifest-overview)
  · [App manifest schema](https://learn.microsoft.com/en-us/microsoft-365/extensibility/schema/)
  · [Requirement configuration](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/understand-requirement-configuration).

Validate against pinned public schema versions and enforce a smaller executable subset. Keep
validation offline and ignore manifest URLs for schema and external entities. Reuse
[`src/core/xml`](../src/core/xml/index.ts) and the bounded ZIP reader in
[`src/core/opc/zip-limits.ts`](../src/core/opc/zip-limits.ts). Initially cap each manifest at 256
KiB, nesting at 32, and metadata packages at 128 entries and 4 MiB inflated. Reject DTDs, external
entities, duplicate JSON keys or resource IDs, unsafe or duplicate archive paths, ambiguous
permission declarations, invalid host mappings, and unresolved required resources. Normalize
HTTPS URLs and origins without credentials. Reject executable URL schemes. Store a normalized
manifest without trusting its declarations to authorize fetching. Treat all labels
as text and defer remote icons. Apply decode and pixel limits if icons are later enabled.

The first UI projects custom tabs and **Show Task Pane** commands onto the existing ribbon
contract. Required unsupported features block activation with a reason. Optional unsupported
features stay disabled with a reason. Function commands, content add-ins, background activation,
shared runtimes, Excel custom functions, and built-in ribbon group contributions need dedicated
slices. A command from a manifest resolves to a host-owned action ID. Catalog data carries no
callbacks into the editor.

## Run document APIs against the live editor

Publish a clean-room browser client that the publisher explicitly includes in its compatible
build in place of the Microsoft CDN script. Keep the client inside the add-in frame. Microsoft
Marketplace certification requires the Microsoft-hosted Office.js for Microsoft Office add-ins.
Support a separate publisher-supplied package for this host. Unmodified Marketplace add-ins that
require Microsoft's CDN library and host protocol have no supported launch path in this design.
[Marketplace certification policy, section 1120](https://learn.microsoft.com/en-us/legal/marketplace/certification-policies#1120-office-add-ins-word-excel-powerpoint-and-outlook).

1. `Office.onReady` resolves with the current host and platform after a validated handshake.
   Implement its promise and callback forms. `Office.context.document` offers the supported
   callback APIs with `AsyncResult` status, values, errors, and caller context. Denied requests
   return errors through the same API. [Initialization](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/initialize-add-in)
   · [Common API model](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/office-javascript-api-object-model).
2. `Word.run`, `Excel.run`, and `PowerPoint.run` create a `RequestContext` for one live document.
   Its root is `document`, `workbook`, or `presentation`. Proxies use typed object paths and
   session-scoped handles. A getter for an unloaded scalar fails with `PropertyNotLoaded`. A
   loaded property reflects the last successful `sync`. Implement used overloads for `load`,
   collection `items`, null-object checks, `ClientResult`, and tracked object lifetime. Preserve
   the documented final implicit `sync` when work is queued. Invalidate untracked handles when
   their `run` ends. [Application-specific API model](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/application-specific-api-model).
3. `context.sync()` sends one ordered batch. Each operation declares its permission, applicable
   host and object types, argument limits, and read or write effect. Resolve paths on the host.
   Check the document and selection generation as needed. Serialize batches per document and
   preserve read-after-write order in the batch. Recheck revision, permissions, read-only state,
   product protection, and preview state before applying edits. Use the normal editor mutation,
   undo, dirty-state, and collaboration paths. Start with atomic batches for supported operations
   and record that host guarantee as a compatibility difference wherever Office reports partial
   effects. Never silently replay a batch after a conflict or a lost response.
4. Put the typed operation dispatcher under `ooxml-core/automation` and keep product operations
   in their core areas. Reuse existing validation and model commands, including Word run edits
   in [`docx.ts`](../src/core/automation/docx.ts), Excel bounds and edit logic in
   [`xlsx-edits.ts`](../src/core/automation/xlsx-edits.ts), and PowerPoint catalogue handlers
   where they expose reusable format operations. Use `createLiveToolHost` as a lifecycle
   reference. Its per-tool save, compare, and reload can discard editor state and costs too much
   for interactive batches. Add a bounded live model adapter per product. Keep frame transport
   and pane DOM in `src/ui`, and bind sessions and policy in `apps/office-suite`. The frame gets
   no MCP registry, file paths, document chooser, or raw editor object.

### Start with published sample workflows

For the compatible Common API client, a host can call
`setOfficeAddInRuntime(root, { origins: ['https://approved-add-ins.example'] })`.
Allow the same execution origin in the embedding site's `frame-src` and provision
publisher-controlled pages without redirects or credentials. The library checks the approved
origin before opening. The opaque sandbox verifies a handshake from the exact frame window
and accepts one private message port. A second frame load closes that port. The parent cannot
authenticate the HTTP origin of an opaque document, so the execution service and CSP must
enforce the host's redirect and origin policy.

In the compatible publisher build, assign
`window.Office = connectOfficeAddIn('https://editor.example')` before the app starts.
Bundle the client into a classic script on the execution origin. The client exposes
`Office.onReady`, `Office.context.host`, and the document's text
`getSelectedDataAsync` and `setSelectedDataAsync` callbacks. Word uses the live text
selection. Excel supports one selected cell. PowerPoint requires a host selection adapter
for the document methods. Other coercions, requirement sets, batch APIs, custom functions,
content add-ins, Microsoft identity and the CDN Office.js host protocol are unsupported.
Client API errors identify unavailable operations and denied permissions. Use
`setOfficeAddInRuntime(root, null)` to disable execution.

Use sample coverage as the initial usage signal. Record which members each sample calls and
count recurring workflows in the pinned sample corpus. Add host-supplied add-in requirements as
they become available. These are implementation targets, with no claim about installed-base
frequency or completed requirement sets. [Microsoft Office Add-in samples](https://learn.microsoft.com/en-us/office/dev/add-ins/overview/office-add-in-code-samples)
· [OfficeDev snippets](https://github.com/OfficeDev/office-js-snippets).

1. **Common APIs.** Start with text coercion for `getSelectedDataAsync` and
   `setSelectedDataAsync` where the host supports each combination. Cover the PowerPoint
   insert-text tutorial and Word and Excel selection samples. Follow with table and matrix
   coercion for Word and Excel, named bindings where available, and settings. Defer
   `AddinCommands 1.1`, which covers automatic
   pane opening. User-initiated ribbon commands have their own manifest support.
   [Common API support](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/support-for-task-pane-and-content-add-ins)
   · [Add-in commands requirement sets](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/common/add-in-commands-requirement-sets)
   · [PowerPoint tutorial](https://learn.microsoft.com/en-us/office/dev/add-ins/tutorials/powerpoint-tutorial-yo).
2. **Excel.** Start with worksheet lookup, selected and addressed range reads and writes,
   formulas, and basic formatting from `ExcelApi 1.1`. Follow with table creation and row
   insertion from the Excel tutorial. Reuse formula and range guards.
   [Excel requirement sets](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/excel/excel-api-requirement-sets).
3. **Word.** Start with selection, body and paragraph reads, insertion, and range formatting
   from `WordApi 1.1`. Add bounded range search and table operations in order of sample demand.
   Preserve unaffected runs, bookmarks, annotations, and section boundaries.
   [Word requirement sets](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets).
4. **PowerPoint.** Start with slide identity and collections from `PowerPointApi 1.2`, then
   shapes and text frames from 1.4. `PowerPointApi 1.1` contains standalone presentation creation,
   so plan that host integration before advertising cumulative sets. Follow with text ranges
   from 1.4, selection from 1.6, custom XML from 1.7, cell border and margin options from 1.8,
   live border and margin objects from 1.9, and shape adjustments from 1.10. Reuse the table
   cell writers in [`core/core/runtime/`](../src/core/pptx/core/core/runtime/), the
   [adjustment validator](../src/core/pptx/core/geometry/preset-adjustment-validation.ts), and the
   existing [custom XML round-trip fixture](../src/core/pptx/__tests__/integration/custom-xml-round-trip.test.ts).
   [PowerPoint requirement sets](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/powerpoint/powerpoint-api-requirement-sets)
   · [Adjustments](https://learn.microsoft.com/en-us/javascript/api/powerpoint/powerpoint.adjustments).

Use one capability registry for manifest admission, the dispatcher, `isSetSupported`, and the
published host-by-member matrix. Compare numeric version components. Report a requirement set
as supported only after its full documented, cumulative contract passes conformance for that
host. Incomplete sets return `false` even when mapped members work. Test those members through
samples with compatible manifests. Show install or activation failures with the missing set,
member, or extension. Match public error codes for implemented error cases and provide a stable
unsupported-operation error for gaps. [Runtime requirement checks](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/specify-api-requirements-runtime).

## Isolate add-ins and enforce consent

Install records availability. Opening requires a host-controlled consent screen that names the
publisher, package version, execution origin, requested data, and possible external destinations.
Each grant covers one add-in identity, approved package and origin, and document session. Enforce
the intersection of the manifest, user grant, host policy, and current document access on every
call. Admin deployment controls availability. Revoking the grant, changing the package or origin,
unmounting the editor, or replacing the document invalidates the channel and pending writes.
Changing active documents never retargets an existing grant. Gate reads of assistant previews
and protected content through the same editor policy as writes.

| Permission                                      | Runtime ceiling                                                                                                                                                                      |
| ----------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| `Restricted` / `Document.Restricted.User`       | The documented restricted subset, such as active view and add-in settings, as implemented. No content or selection reads. Defer settings persistence until storage consent is wired. |
| `ReadDocument` / `Document.Read.User`           | Supported Common API reads, including allowed selection coercions and binding management. Full text file extraction is a documented capability but is deferred.                      |
| `ReadAllDocument` / `Document.ReadAll.User`     | Recognize the broader Common API read scope. Raw OOXML and document bytes remain unavailable during initial rollout.                                                                 |
| `WriteDocument` / `Document.Write.User`         | Supported Common API selection writes. This level does not grant `ReadDocument` access.                                                                                              |
| `ReadWriteDocument` / `Document.ReadWrite.User` | Required for the application-specific APIs even for reads. Permit only implemented methods and authorized effects.                                                                   |
| `ReadWriteMailbox` / `Mailbox.ReadWrite.User`   | Inapplicable to Word, Excel, and PowerPoint. Refuse activation when mailbox access is required.                                                                                      |

Use explicit operation allowlists because the document read and write levels grant different
methods. [Microsoft document permission model](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins)
· [Outlook permissions](https://learn.microsoft.com/en-us/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions).

Run opted-in packages on a host-configured HTTPS origin separate from the suite and its auth
origins. Serve approved static publisher assets and the client from a host-controlled service
that sets the frame's CSP. Start with an opaque-origin iframe using `sandbox="allow-scripts"`.
Omit the flags for same-origin, forms, popups, downloads, and top navigation. Add an iframe
Permissions Policy denying sensors, clipboard, camera, and microphone. Set `referrerpolicy`
to `no-referrer`. Restrict the suite's `frame-src` to configured execution origins. On the
execution origin, start with `default-src 'none'`, `connect-src 'none'`, `object-src 'none'`,
`base-uri 'none'`, `form-action 'none'`, and `frame-src 'none'`. Allow pinned assets through
explicit CSP directives. Keep inline scripts and eval disabled. Set `frame-ancestors` to the
configured host origins. Serve the assets and client without suite credentials or session
cookies. Use a self-contained client script so the opaque frame needs no credentialed module
loads. Require publisher cooperation to meet this deployment model. A parent CSP cannot set the
policy of an arbitrary remote frame. Browser CSP and sandboxing leave some outbound paths,
including self-navigation. Disclose that behavior before handing document data to any add-in.
Hosts requiring zero egress must use a verified
network boundary or decline execution. Record and test per-platform CSP and sandbox behavior.
[Iframe sandbox](https://html.spec.whatwg.org/multipage/iframe-embed-object.html#attr-iframe-sandbox)
· [Content Security Policy](https://www.w3.org/TR/CSP3/).

Use `postMessage` only to establish a versioned, per-open `MessagePort`. The frame sends the
handshake to the exact suite origin. The suite checks `event.source` against the expected iframe,
requires the expected opaque `null` origin, and consumes a one-time nonce tied to that frame
generation. Reject unexpected origins, stale handshakes, and extra ports. Exchange document data
only over the established port. Close it on navigation and teardown, and check the generation
and live grant before returning results or committing writes. Use JSON strings for document
messages. Check encoded size before parsing, then validate exact schemas and bounded depth on
both sides. Reject transferables after the handshake. Start with a 256 KiB encoded request and
response limit, 100 operations per sync, one sync in flight per context, 10 batches per second
per frame, and a 1,000-handle cap per session. Add per-operation limits for range cells, text,
XML, and collection results. Bound pending work and release handles after each run or teardown.
Reject oversized or rate-limited batches before dispatch. Record refusal categories without
payloads. [Web messaging](https://html.spec.whatwg.org/multipage/web-messaging.html).

Initially refuse `displayDialogAsync`, `getAccessToken`, and nested app authentication. Advertise
no dialog or identity requirement sets. A later host-controlled dialog can implement
`displayDialogAsync`, `messageParent`, and close events. Require a user gesture, one dialog per
pane, an initial HTTPS URL on the same origin as the pane, and explicit policy for any redirect.
Use a separate channel with no document capability, bounded messages, and focus restoration.
For auth, a separate opt-in adapter can support a publisher-registered identity flow using
documented OAuth endpoints with state, PKCE, and exact redirect URIs. Gate any Office SSO-shaped
adapter on verified audience, scopes, and a documented provider interface. Keep suite, Microsoft
365, Graph, and host admin tokens out of the add-in channel. [Office dialog API](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/dialog-api-in-office-add-ins)
· [Nested app authentication](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/enable-nested-app-authentication-in-your-add-in)
· [Legacy Office SSO](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/sso-in-office-add-ins).

## Catalogs and Microsoft services

Extend the catalog UI with a trusted provider contract. `list({ host, locale, cursor, signal })`
returns a bounded page of plain records with provider ID, entry ID, version, publisher, label,
supported hosts, source and availability. `install({ id, version, signal })` resolves a manifest
and optional verified package with a content digest and publisher provenance. Restrict remote fetches to configured
origins and revalidate redirects. A server adapter rejects private and local network targets
unless it is explicitly configured for a tenant-private provider. Local sideload
uses the same validation. Verify the previewed identity, version, origin, and requested scopes
against the fetched package before recording installation. Key installs by provider and manifest
identity, and make install and removal idempotent. Keep a validated version usable if an update
fails. Surface permission or origin changes before activation. Distinguish user-removable entries
from required managed assignments and remove a revoked entitlement along with its active grant.

- **BYO and centrally deployed.** A host can supply a static list, manifests, or an authenticated
  catalog on its own backend. For Microsoft 365 tenants, Integrated Apps is the admin portal and
  Centralized Deployment has documented `Get-OrganizationAddIn` and `New-OrganizationAddIn`
  cmdlets, including deployment using a Marketplace asset ID. A tenant admin can use public
  tooling or a manual export to publish authorized assignments to the BYO catalog. The tenant
  backend resolves group membership and supplies only entries for the current user. Microsoft
  admin credentials stay on the admin side. An assignment or product ID alone supplies no
  runnable package. [Deployment options](https://learn.microsoft.com/en-us/office/dev/add-ins/publish/publish)
  · [Centralized Deployment cmdlets](https://learn.microsoft.com/en-us/microsoft-365/enterprise/use-the-centralized-deployment-powershell-cmdlets-to-manage-add-ins).
- **Official Marketplace discovery.** Keep the external Office-filtered listing link. The
  documented Azure Marketplace Catalog `Search` and `Get Product` APIs expose metadata with an
  `Office365` product type covering apps, add-ins, and integrations. An explicit host adapter
  can use these for discovery after verifying coverage for Office offers and the supported
  host mapping. The docs reviewed provide no third-party Word, Excel, or PowerPoint install
  contract for retrieving the executable package or user's store entitlement. The public
  click-and-run links target Microsoft Office. Show them as **Open in Microsoft Office** and
  count them as handoffs. [Catalog Search](https://learn.microsoft.com/en-us/rest/api/marketplacecatalog/dataplane/search/search?view=rest-marketplacecatalog-dataplane-2025-05-01)
  · [Catalog Get Product](https://learn.microsoft.com/en-us/rest/api/marketplacecatalog/dataplane/get-product/get-product?view=rest-marketplacecatalog-dataplane-2025-05-01)
  · [Office installation links](https://learn.microsoft.com/en-us/office/dev/add-ins/publish/publish-office-add-ins-to-appsource#provide-an-installation-link).
- **Publisher distribution.** Partner Center and its Product Ingestion API let authorized
  sellers manage their own offers. Hosts can associate a Marketplace listing with a package
  independently supplied and authorized by that publisher. Identify its install source in the
  UI. Keep native store acquisition open pending a documented or licensed interface for this
  host and a way to run the acquired package. Exclude scraping, private endpoints, and reuse of
  Microsoft's Office host protocol. [Office Add-in Product Ingestion API](https://learn.microsoft.com/en-us/partner-center/marketplace-offers/product-ingestion-office-add-in).

The Store pivot is off by default. Get Add-ins opens My Add-ins until a host enables a
storefront through `setOfficeAddInStorefront`. A host can opt into a Marketplace handoff
or supply a bounded list of offers. The handoff opens only on a user's click, without
sending a referrer or document data. The dialog calls an offer's file reader on selection
and uses the same validation and consent flow as an upload. Admin Managed shows trusted
host ribbon entries.
Remote discovery, account entitlements, publisher verification and tenant synchronization
are future work. Visio exposes the same dialog and host ribbon entries while local manifest
installation remains unavailable.

### Configurable handoff and bring-your-own storefront

A host can register `setOfficeAddInMarketplaceResolver(root, resolver)` independently of
Store. The My Add-ins field accepts an Office listing URL or a WA asset ID. The callback
receives only the normalized asset ID and an abort signal. Pass `null` to disable it.
For example, a host that serves publisher-authorized packages from its own HTTPS backend can
register this resolver:

```js
setOfficeAddInMarketplaceResolver(root, async (assetId, signal) => {
	const response = await fetch(`/office-add-ins/${assetId}`, {
		signal,
		credentials: 'omit',
		redirect: 'error',
		referrerPolicy: 'no-referrer',
	});
	if (!response.ok) throw new Error('Add-in package unavailable');
	const reader = response.body?.getReader();
	if (!reader) throw new Error('Missing package body');
	const result = new Uint8Array(4 * 1024 * 1024);
	let length = 0;
	try {
		for (;;) {
			const { done, value } = await reader.read();
			if (done) return result.slice(0, length);
			if (length + value.byteLength > result.byteLength) throw new Error('Package too large');
			result.set(value, length);
			length += value.byteLength;
		}
	} finally {
		await reader.cancel();
	}
});
```

On the backend, match `/^WA\d{1,16}$/`, look up that exact ID in a fixed registry of
publisher-authorized files, and serve only the registered file after applying the host's
access policy and size limit. Return 404 for unregistered IDs. The asset ID is a lookup key.
The backend needs no URL fetching or Microsoft credentials to implement this reference flow.

One `storefront` option on the add-in host decides what Get Add-ins shows. Future options
include per-listing Office installation links and remote storefront providers.

- **Off (default).** No storefront, no external links and no catalog requests. Installed and
  admin-assigned add-ins still work.
- **Microsoft handoff (opt-in).** Today's Office-filtered Marketplace link can use a host-supplied
  label and HTTPS destination. A future per-listing **Open in Microsoft Office** action could
  use a host-supplied URL template (for example a
  tenant's internal app portal that then deploys through Centralized Deployment). The template
  receives only public listing identifiers, never document, user or tenant data.
- **Bring-your-own storefront.** The host registers a storefront provider and the suite's store
  pane renders it in place of Microsoft's. The provider implements the catalog contract above
  (`list`, `install`) plus optional `search({ query, host, locale, cursor, signal })`,
  `categories({ host, locale })`, `details({ id, version })` (description, screenshots by URL or
  bytes, publisher, permissions, privacy and support links, price or licence text) and
  `entitlement({ id })` for paid or licensed add-ins the provider manages. Several providers can
  be registered at once; the pane shows one tab or section per provider, in the order and with
  the names and icons the host gives, and every entry shows which storefront it came from.
- **Fully custom UI.** A host that wants its own look can replace the store pane with its own
  element or framework component and drive installs through the same `install` call, so
  validation, consent and permission checks stay in the suite.

Whatever the storefront, installation always runs the same path: manifest and package
validation, publisher and digest verification, the consent screen and the permission ceiling.
A storefront can offer an add-in; it cannot skip those checks, grant permissions or reach the
document. Storefront content is untrusted: render descriptions as plain text (or a small
sanitised Markdown subset), load screenshots only from the provider's declared origins (or
provider-supplied bytes), and apply the same request limits as catalogs. Each storefront's
requests, origins and stored data go into `PRIVACY.md` and the suite CSP when it is enabled.
Storefronts, remote providers and execution default to off in the packages.

## Privacy and performance gates

Each shipping slice updates [`PRIVACY.md`](../PRIVACY.md), affected package privacy sections,
and the suite and desktop CSP where applicable. Disclose the following flows before enabling
them. Keep all new remote providers and execution features off by default. Return a local
unavailable status on offline or denied requests and keep the document usable.

| Trigger                                                    | Destination and data                                                                                                                                                                                                                                                                                                                                                 | Off switch                                                                                                                                                  |
| ---------------------------------------------------------- | -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | ----------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Browse Marketplace or open a Microsoft Office install link | Microsoft listing or Office endpoints receive the requested page or public offer ID and title and normal request metadata. Use no opener or referrer.                                                                                                                                                                                                                | Keep the action user-initiated. `storefront: 'off'` or a BYO storefront removes Microsoft links; hosts can relabel or retarget the handoff.                 |
| Open or query an opted-in remote catalog or BYO storefront | Configured BYO endpoint or documented Marketplace adapter receives host, locale, page cursor and optional search text. The BYO endpoint may receive a catalog-scoped credential. Its admin backend may contact Microsoft with tenant assignment data and admin credentials.                                                                                          | Disable the provider or configure a static local catalog. Disable admin synchronization on the backend.                                                     |
| Install or explicitly refresh a remote add-in              | Configured publisher or catalog endpoint receives package or manifest ID and version and normal request metadata. Fetch package-local icons without a network call. Keep remote icons off by default. An enabled remote icon fetch discloses its explicit URL and request metadata.                                                                                  | Remove the install source, use a local package, or disable remote asset fetching. Keep background refresh off.                                              |
| Open a pane or approved dialog                             | Approved execution or dialog origin receives its entry URL and browser request metadata. Browser subresources and any enabled service calls may send publisher-selected data. The bridge can send only data covered by the grant to the frame. The publisher can transmit that data using available browser channels. Resolve approved egress policy before consent. | Deny or revoke the grant, close the pane, or disable runtime, dialogs, or the named provider.                                                               |
| Start a supported sign-in                                  | The configured identity provider and publisher's registered callback receive OAuth request parameters, browser metadata, and authorization results. The configured token endpoint receives the code and verifier.                                                                                                                                                    | Leave auth disabled or disable or disconnect the provider.                                                                                                  |
| Persist installs and optional settings                     | Profile-scoped local storage holds install identity, version, manifest digest, source, and policy metadata. Add-in settings use a namespaced quota and their own consent. Keep document grants and tokens in memory. Document-scoped settings or custom XML written into a file follow its existing save, sharing, and sync paths.                                   | Offer session-only installs, remove add-in data on uninstall, and give hosts a storage-off option. Document-saved data is removed through document editing. |

Use separate lazy imports for the catalog, manifest readers, frame client, and product adapters.
Give the dispatcher and manifest readers a narrow automation subpath so opening a pane does
not pull in the full tool catalogue or another editor's engine.
Fetch no add-in scripts, icons, or catalog entries at editor startup. Load the catalog on Get
Add-ins and the compatible runtime on pane open. Keep dependencies within each product and use
existing core helpers. Set initial design targets of zero added runtime JavaScript bytes on
the default Word, Excel, and PowerPoint routes, at most 20 KiB gzip for the on-demand catalog
UI and manifest path, and at most 40 KiB gzip for the pane client, bridge, and one product
adapter combined. Count incremental first-party chunks including static dependencies. Record
third-party pane transfers separately. Use the existing package and suite byte checks as
release gates. Validate batching with one bridge round-trip per `sync` and zero whole-document
exports, reloads, or scans per routine batch. Cache validated manifests by digest, drop that
cache on update, and bound model queries to the requested selection or range. Move package
validation and large permitted serialization to a worker when practical.
[Repository performance budgets](performance-budgets.md).

## Implementation slices and evidence

| Slice                       | Scope and acceptance tests                                                                                                                                                                                                                                                                                                                                 | Checklist effect                                                                                                                         |
| --------------------------- | ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | ---------------------------------------------------------------------------------------------------------------------------------------- |
| 1. XML install              | Parse XML task panes and overrides. Test official manifest fixtures, locale fallback, malformed input, invalid URLs, requirements and disallowed commands. Drive local install/remove and failed update through the existing dialog.                                                                                                                       | Partial P1 runtime and installation work. The already-checked host catalog item stays complete.                                          |
| 2. Unified install          | Validate JSON and the bounded package. Test paired XML/JSON fixtures, extension requirements, package limits, duplicate keys, hostile paths, and unsupported extension messaging.                                                                                                                                                                          | Partial P1.                                                                                                                              |
| 3. Frame and Common API     | Add consent, isolated bootstrap, and Common API selection operations. Use separate test origins to verify CSP, navigation, wrong source and nonce, cross-document isolation, no early data, revoke races, request budgets, and offline errors. Test keyboard access, named controls, disabled reasons, focus return, and all locale keys.                  | Partial P1.                                                                                                                              |
| 4. Context and Excel        | Land proxy and batch infrastructure with one Excel adapter. Run range read and write samples, unloaded property and stale-handle tests, batching, ordering, protection and formula guards. Save and reopen edited workbooks and verify undo and unrelated parts.                                                                                           | Partial P1 and P2.                                                                                                                       |
| 5. Word                     | Add body, paragraph, and range adapters. Run insertion and formatting samples. Check save and reopen, undo, preservation across run boundaries, and collaboration conflicts.                                                                                                                                                                               | Partial P1 and P2.                                                                                                                       |
| 6. PowerPoint basics        | Wire presentation, slides, selection, and shapes. Run slide and shape samples. Test shape identity after reorder, delete, undo, autosave, and save and reopen.                                                                                                                                                                                             | Partial P1 and P2.                                                                                                                       |
| 7. PowerPoint advanced APIs | Give text ranges, cell borders and margins, shape adjustments, and custom XML one independently testable change each. Run published samples and save/reopen fixtures checking neighboring runs, merged cells, guide values, and unrelated package parts. Exercise protection and limit refusals.                                                           | Closes the named P2 gap after the shared context and all named PowerPoint APIs pass. Report remaining Office.js API coverage separately. |
| 8. Catalog providers        | Implement BYO and managed assignments, the `storefront` option (off, configurable Microsoft handoff, BYO storefront providers, custom store UI) and optional documented Marketplace discovery. Test with mocked providers and a tenant acceptance run where available. Cover paging, revocation, digest mismatch, offline, and publisher-supplied install. | Narrows P1. Native store install stays open until a supported acquisition path passes end to end for all three hosts.                    |
| 9. Dialogs and identity     | Use separate changes for the dialog channel and each provider. Test gesture, origin and redirect checks, messaging, focus, denial, token audience, and revocation. Include public auth samples only for implemented providers.                                                                                                                             | Extends P1 runtime coverage. Does not prove native store acquisition.                                                                    |
| 10. Forms provider          | In a separate effort, validate BYO response delivery with a tenant flow and narrowly scoped workbook updates. Track native PowerPoint embedding, management, and Microsoft-hosted workbook sync against their own acceptance criteria.                                                                                                                     | Partial adjacent Forms P1 only.                                                                                                          |

Build the conformance harness with slice 3. Run pinned, license-compatible Microsoft quickstart
and Script Lab samples in the same isolated frame as add-ins. Record source URL, commit,
license, host, exercised members, fixture, and expected results. Preserve the sample logic and
document bootstrap-only substitutions. Run CI with external network denied, reproducible
documents, and a fake catalog. Compare callback results, proxy behavior, and saved document
changes. For gaps in the public docs, run the same published sample in Microsoft Office with
synthetic data and record observable behavior. Keep unsupported samples visible with the
specific missing capability. Add a security corpus case when fixing a defect class. Run the
focused Vitest files, strict and PowerPoint typechecks, core and UI builds, byte checks, and
per-product browser acceptance for each implementation slice. The two open add-in checklist
items remain open after this design.
