# Migrating to OOXML 2.0

OOXML 2.0 is one coordinated major release of every package in this repository. This guide lists every change in it that can break your code, your configuration, your tests or your stored data, with what to write instead.

The sections go from the packages most people use to the ones few people run. Within a section, the change that reaches the most users comes first. Read the sections for the packages you install, then [Before you start](#before-you-start), which applies to everyone. The full notes of each package, breaking changes included, are at the top of its `CHANGELOG.md` and in its GitHub release.

- [Before you start](#before-you-start)
- [PowerPoint viewers](#powerpoint-viewers)
- [Word, Excel and Visio viewers](#word-excel-and-visio-viewers)
- [Themes and shared elements in `ooxml-ui`](#themes-and-shared-elements-in-ooxml-ui)
- [PowerPoint building blocks in `ooxml-ui/pptx`](#powerpoint-building-blocks-in-ooxml-uipptx)
- [`pptx-viewer-core` and `ooxml-core/pptx`](#pptx-viewer-core-and-ooxml-corepptx)
- [Shared areas of `ooxml-core`](#shared-areas-of-ooxml-core)
- [MCP packages and `ooxml-core/automation`](#mcp-packages-and-ooxml-coreautomation)
- [OpenTeams](#openteams)
- [OOXML Office deployments](#ooxml-office-deployments)

## Before you start

### Upgrade every OOXML package at once

Each package moves to its own next major version: `ooxml-core` and `ooxml-ui` to 2, the PowerPoint viewers to their next major, and so on. The viewers are tested only against the `ooxml-core` and `ooxml-ui` released with them, so move every OOXML package in one change.

Before:

```json
{
	"dependencies": {
		"ooxml-core": "^1.54.1",
		"pptx-react-viewer": "^4.81.0"
	}
}
```

After:

```json
{
	"dependencies": {
		"ooxml-core": "^2.0.0",
		"pptx-react-viewer": "^5.0.0"
	}
}
```

Run `npm install <package>@latest` for each OOXML package you list, then check that your lockfile holds one version of `ooxml-core` and one of `ooxml-ui`.

### Check the browsers you support

The packages support the last two major versions of Chrome, Edge, Firefox and Safari, and Safari 16.4 at least. [browser-support.md](browser-support.md) lists the Safari behaviours they handle. If you support an older browser, test it before you upgrade.

### Deploy the new decoder folders if you copy files by hand

`emf-converter` 5 loads its image decoders and colour data on demand. Package managers and bundlers need no change. If you copy compiled files of `ooxml-core`, `ooxml-ui`, `pptx-viewer-core` or a viewer by hand, copy the chunk folders too and keep their relative paths.

Before:

```text
public/vendor/ooxml-core/index.mjs
```

After:

```text
public/vendor/ooxml-core/index.mjs
public/vendor/ooxml-core/browser-chunks/
public/vendor/ooxml-core/node-chunks/
```

## PowerPoint viewers

This section covers `pptx-react-viewer`, `pptx-vue-viewer`, `pptx-angular-viewer`, `pptx-svelte-viewer`, `pptx-vanilla-viewer` and `@christophervr/pptx-viewer`.

### Linked pictures, media and sounds wait for Enable Content

As in Microsoft PowerPoint, a deck's linked pictures, fills, backgrounds, bullets, audio, video, online video and linked animation, transition and hyperlink sounds are no longer downloaded when it opens. A security bar offers **Enable Content**, which redraws the deck in place: nothing reloads and undo history is kept. Each participant in a collaboration session decides for themselves.

To keep downloading linked content without asking, or to block it without the bar:

```tsx
<PowerPointViewer source={file} allowExternalContent={true} />  // download, no bar
<PowerPointViewer source={file} allowExternalContent={false} /> // block, no bar
```

A custom shell built from building blocks renders the bar from `externalContent` (or the binding's `ExternalContentBar` component) and passes `externalContent.allowed` as `linkedContent` to slides it draws itself: `SlideThumbnail`, `renderSlideStage`, `renderSlideThumbnail` and `createSlidesPane`. A surface that is never told the consent withholds linked content.

A host that draws slides from the viewer's model must read linked addresses through `linkedContentView(value, allowed)` from `ooxml-ui/pptx`, because the model now keeps them.

`applySlideTransitionSound` takes the viewer's media map and consent, and `TransitionSoundPlayer` is removed:

```ts
// Before
applySlideTransitionSound(transition, resolveUrl, player);
// After
applySlideTransitionSound(transition, { mediaUrls, linkedContent: allowed });
```

### Tab reaches the slide before its objects

Tab no longer selects slide objects from anywhere in the viewer. The slide is one Tab stop. Press Enter on it, reach it with F6, or click it, and then Tab and Shift+Tab walk its objects as before. In the speaker notes of the React, Vue and Angular viewers, Tab and Shift+Tab now move focus instead of indenting, as in the vanilla and Svelte viewers.

End-to-end tests that pressed Tab straight after loading a deck must enter the slide first.

Before:

```ts
await page.keyboard.press('Tab'); // selected the first object
await page.keyboard.press('Delete');
```

After:

```ts
await page.locator('[data-pptx-viewport]').click(); // or press F6 until the slide has focus
await page.keyboard.press('Enter');
await page.keyboard.press('Tab'); // selects the first object
await page.keyboard.press('Delete');
```

Tell users who indent notes with Tab to use the notes toolbar's **Increase Indent** and **Decrease Indent** buttons.

### Links in a deck always open in a new tab

External links open in a new tab with `noopener` and `noreferrer`, whatever target frame the deck names. The target frame is still saved back unchanged. `file:` and `ftp:` links are no longer opened. Relative links, such as `notes.html` or `#part`, are no longer opened unless the host opts in.

If your decks link to pages beside the host page, opt in once at startup. The setting applies to every viewer on the page, and the links resolve against the host page's address.

Before:

```ts
import { PowerPointViewer } from 'pptx-react-viewer';
```

After:

```ts
import { setOpenRelativeHyperlinks } from 'ooxml-ui/pptx';
import { PowerPointViewer } from 'pptx-react-viewer';

setOpenRelativeHyperlinks(true);
```

### Peer-to-peer Share and Broadcast links need a session key

A Share or Broadcast session without a server is now keyed by a random secret, which its link carries after `#key=`. The public signaling server sees only a hash of it. Links made by an earlier version, and bare room codes without a server, no longer join. Start a new session and send its new link.

A `CollaborationConfig` with `transport: 'webrtc'`, or with a blank `serverUrl`, must carry `secret`. `authToken` is no longer the peer-to-peer password.

Before:

```ts
const collaboration: CollaborationConfig = {
	roomId: 'deck-42',
	userName: 'Alice',
	serverUrl: '',
	transport: 'webrtc',
	authToken: 'shared-password',
};
```

After:

```ts
import { buildCreateCollaborationConfig } from 'ooxml-ui/pptx';

const collaboration = buildCreateCollaborationConfig({
	roomId: 'deck-42',
	userName: 'Alice',
	serverUrl: '',
}); // adds `transport: 'webrtc'` and a fresh `secret`
```

To join from a link your own code received, read it with `readShareInvitation` and pass the secret on:

```ts
import { readShareInvitation } from 'ooxml-core/collab';

const invitation = readShareInvitation(link);
if (invitation.kind === 'session') {
	const { roomId, serverUrl, transport, secret } = invitation.target;
	join({ roomId, serverUrl, transport, secret, userName: 'Ben', sessionIntent: 'join' });
} else if (invitation.kind === 'missing-secret') {
	showMessage('This link has no session key. Ask for a new link.');
}
```

`buildShareUrl` and `buildQueryLinkUrl` are removed from `ooxml-ui/pptx` and `pptx-angular-viewer`, and `buildBroadcastViewerUrl` takes the session's config.

Before:

```ts
const shareLink = buildShareUrl(roomId, serverUrl, window.location);
const viewerLink = buildBroadcastViewerUrl(roomId, serverUrl, window.location);
```

After:

```ts
import { buildBroadcastViewerUrl, buildCollaborationShareUrl } from 'ooxml-ui/pptx';

const shareLink = buildCollaborationShareUrl(config, window.location);
const viewerLink = buildBroadcastViewerUrl(config, window.location);
```

### Theme and presentation tools follow the assistant's write policy

Theme and presentation-level assistant tools now follow `writePolicy`. Under `stage` and `approve` they create proposals instead of applying at once. Under `approve`, each change must be accepted on its own, and **Accept all** skips those changes. `delete_slides` and `import_from_json` behave this way under every policy.

If your host relied on theme or metadata tools applying at once, choose `auto`.

Before:

```ts
const ai = createAiAssistant({ connection: { kind: 'endpoint', api: '/api/ai/chat' } });
```

After:

```ts
const ai = createAiAssistant({
	connection: { kind: 'endpoint', api: '/api/ai/chat' },
	writePolicy: 'auto',
});
```

The default stays `stage`. Keep it if your users review every change.

### Generated SVG ids are encoded

The ids of generated SVG filters, gradients, markers, patterns and clip paths are now encoded, so most of them change even for ordinary decks. Styles or code that look one up by id must get it from the helper for that effect.

Before:

```ts
const filter = svg.querySelector(`#soft-edge-${element.id}`);
```

After:

```ts
import { getSoftEdgeFilterId } from 'ooxml-ui/pptx';

const filter = svg.getElementById(getSoftEdgeFilterId(element.id));
```

For an effect with no helper, encode the element id with `svgIdToken` from `ooxml-ui/pptx`. For example, `soft-edge-ppt/slides/slide1.xml-shape-0` is now `soft-edge-ppt_2f_slides_2f_slide1.xml-shape-0`.

### Collaboration state carries the reason a session failed

`CollaborationSessionState` and `CollaborationShellState` gain a required `failure` field, and the React `StartSessionForm` prop `connectionFailed` becomes `connectionFailure`, a translation key. This matters only to hosts that build these objects or render that form themselves.

Before:

```tsx
<StartSessionForm {...props} connectionFailed={state.status === 'error'} />
```

After:

```tsx
<StartSessionForm
	{...props}
	connectionFailure={shareFormFailureKey(
		{ mode, invitation },
		collab ? collaborationFailureMessageKey(collab.status, collab.failure) : null,
	)}
/>
```

`shareFormFailureKey` and `collaborationFailureMessageKey` are exported from `ooxml-ui/pptx`.

### The leave message must come from the embedding page

A viewer in a frame obeys the collaboration leave `postMessage` only when the embedding window, or the frame itself, sends it from the frame's own origin. If your page embeds a viewer from another origin and posts the leave message, name your page's origin.

Before:

```ts
const collaboration: CollaborationConfig = { roomId, userName, serverUrl };
```

After:

```ts
const collaboration: CollaborationConfig = {
	roomId,
	userName,
	serverUrl,
	leaveMessageOrigin: 'https://host.example.com',
};
```

### Tests find the zoom button by its new name

The status bar's zoom percentage button is now named by the percentage it shows first, for example "100%, Zoom to fit". This also applies to `office-ui-status-bar` in Word, Excel and Visio.

Before:

```ts
await page.getByRole('button', { name: 'Zoom to fit' }).click();
```

After:

```ts
await page.locator('[data-id="zoomFit"]').click();
```

Hosts that translate the **Design** > **Fonts** tiles add the keys `pptx.gallery.themeFonts.schemeName` and `pptx.gallery.themeFonts.currentName`, which now give the tile names. `pptx.gallery.themeFonts.scheme` and `pptx.gallery.themeFonts.current` give only the tooltips.

## Word, Excel and Visio viewers

This section covers the `docx-*-viewer`, `xlsx-*-viewer` and `visio-*-viewer` packages, `@christophervr/xlsx-react-viewer`, and the editors in `ooxml-ui/docx`, `ooxml-ui/xlsx` and `ooxml-ui/visio`.

### The editors take dark mode from the theme, not from `color-scheme`

The Word, Excel and Visio editors no longer inherit the page's CSS `color-scheme`. They read the new `--office-color-scheme` token, which is `light` unless the theme's dark mode applies. A container that sets only `color-scheme: dark` now gets light editors.

Before:

```css
.workspace {
	color-scheme: dark;
}
```

After:

```css
.workspace {
	color-scheme: dark;
	--office-color-scheme: dark;
}
```

You can also put `data-office-theme="dark"` on the container, which sets both.

### Choose `surfaceOnly` before an Excel editor connects

Changing `surfaceOnly` on an `xlsx-editor` that is already attached to the page now throws, as it does on `docx-editor`.

Before:

```ts
const editor = document.createElement('xlsx-editor');
host.append(editor);
editor.surfaceOnly = true;
```

After:

```ts
const editor = document.createElement('xlsx-editor');
editor.surfaceOnly = true;
host.append(editor);
```

### An editor removed from the page leaves its room

A `docx-editor` or `xlsx-editor` that shares a room now leaves it when it is removed from the page and not attached again before the current task ends. It joins the same room again when it is attached later. Moving an editor between parents in one task keeps it in the room. If you detach an editor to hide it and expect it to keep syncing, hide it with CSS instead.

Before:

```ts
panel.removeChild(editor); // stayed in the room
setTimeout(() => otherPanel.append(editor), 500);
```

After:

```ts
editor.hidden = true; // stays attached and in the room
```

## Themes and shared elements in `ooxml-ui`

### Derived tokens resolve where they are read

The theme no longer declares derived tokens, the tokens whose default reads another token, on `:root` or on the `data-office-theme` scopes. Examples are `--office-accent-ink`, `--office-popover`, `--office-field-border` and `--office-select-background`. Each control resolves them where it reads them, so a container that overrides `--office-accent` now also changes the accent text inside it. `themeBridge` and `shadcnBridge` no longer repeat those defaults, and `isDerivedToken` tells you which tokens are derived.

Code that reads a derived token from computed style gets an empty string unless a host set it.

Before:

```ts
const popover = getComputedStyle(document.documentElement).getPropertyValue('--office-popover');
```

After:

```ts
import { tok } from 'ooxml-ui';

// In CSS you write yourself, read the token with its default:
const rule = `.my-panel { background: ${tok('--office-popover')}; }`;
```

To compute the value in script, resolve the default from `OFFICE_TOKENS`, which `ooxml-ui` also exports.

### Form controls use `--office-control-border`

Form control boundaries now use the new `--office-control-border` token, `#8a8886` by default, so they reach 3:1 against the background. The product `input` colours (`--dve-input`, `--xve-input`, `--pptx-input`) default to it and colour the shared form controls inside the editor. A host that set `--office-border` to restyle fields sets the new token.

Before:

```css
.host {
	--office-border: #c8c6c4;
}
```

After:

```css
.host {
	--office-border: #c8c6c4;
	--office-control-border: #8a8886; /* keep it at 3:1 against the background */
}
```

Hosts that already set `--office-field-border`, `--office-checkbox-border` or `--office-switch-border` keep their values.

### Share dialog labels need `missingKey`

`OfficeShareDialogLabels` gains a required `missingKey` caption. This matters only to code that builds a complete labels object.

Before:

```ts
const labels: OfficeShareDialogLabels = { ...myLabels };
```

After:

```ts
const labels: OfficeShareDialogLabels = {
	...myLabels,
	missingKey: 'This link has no session key. Ask for a new link.',
};
```

## PowerPoint building blocks in `ooxml-ui/pptx`

This section is for hosts that build their own PowerPoint shell from `ooxml-ui/pptx` and `ooxml-ui/pptx/ai`.

### The keymap takes `slideFocus` and has an `editText` action

`isEditorControlTarget` and the `isControlTarget` flag of `EditorKeyGuard` are gone. The keymap now takes `slideFocus` (`'outside' | 'pane' | 'objects'`), and `slideKeyFocus(target)` computes it. `dispatchEditorKey` fills it in when you omit it.

Before:

```ts
const result = mapEditorKey(event, {
	...guard,
	isControlTarget: isEditorControlTarget(event.target),
});
```

After:

```ts
import { mapEditorKey, slideKeyFocus } from 'ooxml-ui/pptx';

const result = mapEditorKey(event, { ...guard, slideFocus: slideKeyFocus(event.target) });
```

The keymap has a new action, `editText`, for Enter on a selected object. A `switch` over every `EditorKeyActionName` needs a case for it, and `EditorKeyHandlers` takes an `editText` handler.

Before:

```ts
dispatchEditorKey(event, state, { deleteSelection, nudge });
```

After:

```ts
import { dispatchEditorKey, editSlideObjectText } from 'ooxml-ui/pptx';

dispatchEditorKey(event, state, {
	deleteSelection,
	nudge,
	editText: () => editSlideObjectText(document, selectedId),
});
```

To let Tab walk the objects of your own surface, mark its container with `data-pptx-viewport` and attach `attachViewportNavigation`.

### Hyperlink helpers lose their target argument

`resolveHyperlinkTargetAttrs` is removed, and `safeOpenUrl` and `openExternalHyperlinkConfirmed` no longer take a target argument.

Before:

```ts
const { target, rel } = resolveHyperlinkTargetAttrs(link.tgtFrame);
safeOpenUrl(link.url, target);
await openExternalHyperlinkConfirmed(options, link.url, context, target);
```

After:

```ts
const target = '_blank';
const rel = 'noopener noreferrer';
safeOpenUrl(link.url);
await openExternalHyperlinkConfirmed(options, link.url, context);
```

### Embedded payloads open only as inert data URLs

`isBrowserOpenableMime(mime)` is replaced by `isBrowserOpenableDataUrl(url)`, which reads the type from the payload itself. `openUrlInNewTab` now opens only a `data:` URL of an inert type (PDF, plain text, CSV, JSON, PNG, JPEG, GIF, WebP, BMP, AVIF) and returns whether it opened a tab. It refuses `http(s)`, `blob:` and markup URLs. `downloadDataUrl` ignores anything but `data:` and `blob:` URLs.

Before:

```ts
if (isBrowserOpenableMime(element.oleEmbeddedMimeType)) {
	openUrlInNewTab(element.oleEmbeddedData);
} else {
	downloadDataUrl(element.oleEmbeddedData, fileName);
}
```

After:

```ts
import { openOrDownloadPayload } from 'ooxml-ui/pptx';

openOrDownloadPayload(element.oleEmbeddedData, fileName); // downloads what a tab may not open
```

### Custom assistant bridges report the deck load

`PptxAiBridge` and `PptxAiBridgePort` require `getDeckLoad()`. Staged proposals belong to one presentation load, and they are dropped when the number changes. Return a new number each time the host loads a presentation: open, reopen, **File** > **New**, a recovery restore, or a collaboration join whose room replaces the slides. Return the same number through every edit, including edits that replace the handler.

Before:

```ts
const port: PptxAiBridgePort = {
	getSlides: () => store.slides,
	getActiveSlideIndex: () => store.activeSlide,
	// ...the other members
};
```

After:

```ts
let deckLoad = 0;
function onPresentationLoaded() {
	deckLoad += 1;
}

const port: PptxAiBridgePort = {
	getSlides: () => store.slides,
	getActiveSlideIndex: () => store.activeSlide,
	getDeckLoad: () => deckLoad,
	// ...the other members
};
```

`checkPptxAiBridgeReadsItsWrites(bridge)` reports whether a custom bridge returns its own writes from the next read, which **Apply all** relies on.

The other assistant changes in `ooxml-ui/pptx/ai`:

- `StagedProposal` no longer has an `updater`.
- `buildToolExecutors` takes a run limiter as its fourth argument instead of an animator.
- `session.transport` already adds the deck and focus context, so remove the context your binding or host added itself.
- A custom review list reads `requiresApproval` and `conflict` on `ProposalView`. `proposals.apply(id)` returns `false` for a conflict and keeps the proposal listed.
- Show **Accept all** only when `canAcceptAllProposals(proposals)` is true.
- When `proposals.discarded` is true, say why the list emptied, for example with the `pptx.ai.proposalsDiscarded` string.

Before:

```ts
acceptAllButton.hidden = proposals.length === 0;
```

After:

```ts
import { canAcceptAllProposals } from 'ooxml-ui/pptx/ai';

acceptAllButton.hidden = !canAcceptAllProposals(proposals);
```

## `pptx-viewer-core` and `ooxml-core/pptx`

### Signature verification

Signature verification now reports trust only for a signature that is intact, covers the document content and chains to a root you configured. The changes below are in `ooxml-core/pptx/signature-node`, which `pptx-viewer-core` re-exports as `pptx-viewer-core/signature-node`.

No trust roots are built in, and the platform's TLS roots are no longer used. An intact signature reports `verified-untrusted` until you configure the roots you trust for document signing.

Before:

```sh
node verify-deck.mjs deck.pptx
```

After:

```sh
PPTX_VIEWER_TRUST_ROOTS_FILE=/etc/ooxml/document-signing-roots.pem node verify-deck.mjs deck.pptx
```

`PPTX_VIEWER_TRUST_ROOTS_PEM` takes the PEM text itself.

A new status, `coverage-incomplete`, is reported when a signature leaves stored content unsigned. It appears on the report and on each signature detail, and `SignatureDetail` gains `uncoveredParts` and `coverageGaps`.

Before:

```ts
const report = await inspectPptxDigitalSignatures(bytes);
if (report.verificationStatus !== 'verified-trusted') warn(report.verificationStatus);
```

After:

```ts
const report = await inspectPptxDigitalSignatures(bytes);
if (report.verificationStatus === 'coverage-incomplete') {
	const parts = report.details?.flatMap((detail) => detail.uncoveredParts) ?? [];
	warn(`The signature does not cover ${parts.join(', ')}`);
} else if (report.verificationStatus !== 'verified-trusted') {
	warn(report.verificationStatus);
}
```

`SignatureReferenceCheck.digestStatus` can now be `content-type-mismatch`. `referenceChecks` is empty unless the signature value verified, and `signingTime` is read only from signed content.

`signPptxWithCertificate` now returns `success: false` and no `signedData` for a stored item whose name is not a valid part name, for a relationships part the strict reader rejects, and for output whose signature does not verify and cover its content. Check `success` before you use the result, and rename such items before signing.

Before:

```ts
const result = await signPptxWithCertificate(bytes, certificate, options);
await writeFile('signed.pptx', result.signedData!);
```

After:

```ts
const result = await signPptxWithCertificate(bytes, certificate, options);
if (!result.success || !result.signedData) throw new Error(result.error ?? 'Signing failed');
await writeFile('signed.pptx', result.signedData);
```

The removed helpers have these replacements.

Before:

```ts
const status = verifySignatureValue(signatureXml, certificates);
const checks = buildReferenceChecksFromSignatureXml(signatureXml /* ... */);
```

After:

```ts
import {
	checkSignedReferences,
	parseXmlSignature,
	verifySignedInfo,
} from 'ooxml-core/pptx/signature-node';

const parsed = parseXmlSignature(signatureXml);
if (!parsed.ok) throw new Error(parsed.error);
const verification = verifySignedInfo(parsed.signature);
const references =
	verification.status === 'verified'
		? await checkSignedReferences(
				parsed.signature,
				verification.canonicalSignedInfo,
				signaturePackage,
			)
		: undefined;
```

`signaturePackage` is the `SignaturePackage` of the deck being checked. `verifySignedInfo` reads the signer certificate from the signature itself.

The remaining signature changes need a decision more than code:

- Certificates whose extended key usage excludes document signing, such as TLS server certificates and certificates that list only `clientAuth` or only `codeSigning`, report as untrusted. So do certificates whose key usage has neither `digitalSignature` nor `nonRepudiation`, and chains that exceed a `pathLenConstraint`. Reissue such signing certificates if their signatures must report `verified-trusted`.
- `PPTX_VIEWER_REQUIRE_TIMESTAMP` can no longer be met, because timestamp tokens are not validated. With it set, every signature reports `timestamp-untrusted`. Remove it unless you want that.
- `computeDetailStatus` and `computeVerificationStatus` decide integrity first. A failed signature value, reference or coverage check now outranks certificate, revocation, timestamp and policy statuses, so some packages report a different status than before.
- `PptxHandler.load` rejects a signed package when a relationships part has an element outside the OPC namespace, an exactly repeated `Id`, an invalid `TargetMode` or attributes the schema does not allow. Unsigned packages load as before.
- `resolveReferenceUriToPart` and `parseReferenceUri` strip the `?ContentType=` query, keep the percent-encoding of the part name, and return `undefined` for a URI that does not name exactly one valid part name. The enveloped signature transform is no longer supported.
- `verifySignatureValue`, `buildReferenceChecksFromSignatureXml`, `buildReferenceChecksFromPptxViewerManifest`, `PPTX_VIEWER_MANIFEST_NS` and `normalizePartPath` are removed. `evaluateSignatureCoverage` no longer takes the signature paths, and `SignaturePackage` gains `namesByKey`.
- `parseRelationshipsPart` moved to `ooxml-core/opc` and throws on a part the schema does not allow.

### Revocation checks use `node:http`

OCSP revocation checks now use `node:http` and `node:https` instead of the global `fetch`, and connect only to globally routable addresses. A replaced `globalThis.fetch`, a fetch dispatcher or a proxy configured for fetch no longer sees these requests. Allow direct outbound HTTP and HTTPS to public responders, intercept at the `node:http` level in tests, or turn the check off.

Before:

```ts
globalThis.fetch = mockOcspFetch;
const report = await inspectPptxDigitalSignatures(bytes);
```

After:

```ts
const report = await inspectPptxDigitalSignatures(bytes, { checkRevocation: false });
// The revocation status is `not-checked`, and no request is made.
```

### Element ids encode unusual part names

Element ids are now built from an encoded form of the part name. Decks whose part names use only letters, digits, `-`, `.`, `_` and `~` keep exactly the same ids, such as `ppt/slides/slide1.xml-shape-0`. A part name with any other character, for example a space or a non-ASCII letter, appears percent-encoded in the id. `PptxSlide.id` is unchanged and is still the part name.

If you stored element ids of such decks, reload the deck and read the ids again.

Before:

```ts
const shapeId = 'ppt/slides/my slide.xml-shape-0';
```

After:

```ts
const shapeId = slide.elements[0].id; // 'ppt/slides/my%20slide.xml-shape-0'
```

Autosaved decks reopen normally. An unsaved inline text draft of such a deck is not reapplied, because its element id no longer matches. Existing collaboration rooms keep the ids already stored in them.

### Slides read from a collaboration room are checked

`readSlidesFromYDoc` and `readSlideFromYMap` now check what a peer wrote. An element id that is not in safe token form is encoded as the loader encodes part names. A picture, media, model or embedded-file field keeps only a `data:` URL of its kind, plus `http(s)` URLs or package paths where the loader itself puts them, and a MIME type field must be well formed. Anything else is left out of that participant's view and reported as a `COLLABORATION_REMOTE_VALUE_REJECTED` slide warning.

Before:

```ts
const slides = readSlidesFromYDoc(doc);
```

After:

```ts
const slides = readSlidesFromYDoc(doc);
const rejected = slides.flatMap((slide) =>
	(slide.warnings ?? []).filter((w) => w.code === 'COLLABORATION_REMOTE_VALUE_REJECTED'),
);
```

The viewers show these warnings as a toast. `createRemoteValueNotifier` and `remoteValueWarnings` in `ooxml-ui/pptx` do the same for a custom shell.

### `embeddedWorkbookData` is set only when a load needs it

`PptxChartData.embeddedWorkbookData`, the parsed rows of a chart's embedded workbook, is now set only when the chart part caches no categories, no series values or no scatter x values, or when a date axis has no `c:date1904`. To read the rows of any chart, parse the workbook bytes, which are still set.

Before:

```ts
const rows = chartData.embeddedWorkbookData;
```

After:

```ts
import { parseEmbeddedXlsx } from 'pptx-viewer-core';

const bytes = chartData.externalData?.embeddedWorkbookData;
const rows = chartData.embeddedWorkbookData ?? (bytes ? await parseEmbeddedXlsx(bytes) : undefined);
```

## Shared areas of `ooxml-core`

### Collaboration updates are bounded, and a refusal stops the sync

`ooxml-core/collab` refuses an update that holds more than one struct per 4 bytes, or more than one deleted range per 2 bytes, before it decodes it. The budget grows with the update, so the first sync of a real document stays under half of it. Every transport provider and the Teams server apply it. `DEFAULT_UPDATE_LIMITS` is replaced by `defaultUpdateLimits(bytes)`, `MIN_BYTES_PER_STRUCT` and `MIN_BYTES_PER_DELETE_RANGE`.

Before:

```ts
const result = applyUpdateSafe(doc, update, origin, DEFAULT_UPDATE_LIMITS);
```

After:

```ts
const result = applyUpdateSafe(doc, update, origin); // proportional limits by default
// or set absolute limits:
const strict = applyUpdateSafe(doc, update, origin, {
	maxStructs: 200_000,
	maxDeleteRanges: 100_000,
});
```

A refused update now stops the sync. The provider disconnects with status `error` and emits an `UpdateRefusedError` with a `code` (`too-large`, `malformed`, `server-rejected` or `server-full`). The session reports it as `refusal`, and `collaborationStateOf` includes it. A server close with code 1009 stops the WebSocket transport at once. `ApplyResult` failures carry a `code`, and the `onRejected` callback of `guardProviderUpdates` receives the `UpdateRefusedError` instead of a string.

Before:

```ts
guardProviderUpdates(provider, {
	onRejected: (reason) => {
		if (reason.includes('too large')) showError(reason);
	},
});
```

After:

```ts
import { isUpdateRefusal } from 'ooxml-core/collab';
import { translateCollabText } from 'ooxml-core/i18n';

guardProviderUpdates(provider, {
	onRejected: (error) => {
		if (isUpdateRefusal(error) && error.code === 'too-large') {
			showError(translateCollabText(error.message, navigator.language));
		}
	},
});
```

### `readShareInvitation` replaces `parseShareInvitation`

`parseShareInvitation` returned a target or `null`. `readShareInvitation` returns `{ kind: 'session', target }`, `{ kind: 'missing-secret', roomId }` or `{ kind: 'none' }`.

Before:

```ts
const target = parseShareInvitation(link);
if (target) join(target);
```

After:

```ts
const invitation = readShareInvitation(link);
if (invitation.kind === 'session') join(invitation.target);
```

A `'missing-secret'` result is a peer-to-peer link without its key. Tell the user to ask for a new link.

### The leave message needs the embedding page's origin

`createCollabSession` obeys the leave `postMessage` only from the frame's own origin unless you list the embedding page's origin.

Before:

```ts
createCollabSession({ ...options, teardown: { leaveMessage: 'ooxml:leave' } });
```

After:

```ts
createCollabSession({
	...options,
	teardown: { leaveMessage: 'ooxml:leave', leaveMessageOrigin: 'https://host.example.com' },
});
```

For PowerPoint, set `CollaborationConfig.leaveMessageOrigin`, as shown under [PowerPoint viewers](#the-leave-message-must-come-from-the-embedding-page).

### Hyperlinks keep every safe scheme

The hyperlink policy in `ooxml-core/opc` now stores every well-formed scheme that is not script-capable, such as `tel:`, `sip:`, `msteams:` and `onenote:`. It still opens only http, https and mailto unless a host opts in. The `unsupported-scheme` rejection reason is gone.

Before:

```ts
switch (rejection.reason) {
	case 'unsupported-scheme':
		return t('link.unsupported');
	default:
		return t('link.blocked');
}
```

After:

```ts
switch (rejection.reason) {
	default:
		return t('link.blocked');
}
```

### `diffOpcPackages` has limits

`diffOpcPackages` rejects a package over the default limits of 100 MiB of input, 100 MiB inflated and 10,000 parts. Pass larger limits as the third argument, or handle the rejection.

Before:

```ts
const diff = await diffOpcPackages(base, other);
```

After:

```ts
const diff = await diffOpcPackages(base, other, {
	maxInputBytes: 400 * 1024 * 1024,
	maxEntryBytes: 400 * 1024 * 1024,
	maxTotalBytes: 400 * 1024 * 1024,
});
```

### Chart workbook writes list changed cells

`ChartWorkbookWrite` in `ooxml-core/chart` now lists the cells to change. `cells` holds `{ index, value }` entries, where an empty value clears the cell, and `cellCount` is the size of the range. It replaces `values`.

Before:

```ts
for (const write of writes) writeRange(write.ref, write.values);
```

After:

```ts
for (const write of writes) {
	for (const { index, value } of write.cells) writeCell(write.ref, index, value);
}
```

To create a write from a whole range, build `cells` with `chartWorkbookCells(values)` and set `cellCount` to the length of the range.

`chartSpaceWorkbookWrites` and `chartExSpaceWorkbookWrites` no longer write a chart's blank points. Given the chart as loaded, they return only the cells the edit changed.

Before:

```ts
const writes = chartSpaceWorkbookWrites(edited);
```

After:

```ts
const writes = chartSpaceWorkbookWrites(edited, loaded); // only the cells this edit changed
```

`addChartCategory` no longer gives a category to a series that has none, and it puts the new category in the row after all the chart's data.

### Word style catalogs are values

`ooxml-core/docx` and `docx-core` now resolve style inheritance once per catalog `styles` record. Change a style by replacing the record and the style, not by changing a style object in place.

Before:

```ts
model.paragraphStyles!.styles.Heading1!.basedOn = 'Title';
```

After:

```ts
const catalog = model.paragraphStyles!;
const heading = catalog.styles.Heading1!;
model = {
	...model,
	paragraphStyles: {
		...catalog,
		styles: { ...catalog.styles, Heading1: { ...heading, basedOn: 'Title' } },
	},
};
```

The same holds for `characterStyles` and `tableStyles`.

## MCP packages and `ooxml-core/automation`

This section covers `ooxml-mcp`, `docx-viewer-mcp`, `xlsx-viewer-mcp`, `visio-viewer-mcp`, `pptx-viewer-mcp` and `ooxml-core/automation`.

### Formulas that reach outside the workbook are refused

`xlsx_set_cells`, `xlsx_propose_edits` and `xlsx_apply_proposal` (now schema version 2), and `setXlsxCells`, `applyXlsxEdits`, `proposeXlsxEdits`, `applyXlsxProposal` and `createXlsxFromSheets`, refuse a formula that can reach outside the workbook. This covers web and external data functions such as `WEBSERVICE`, `IMAGE`, `HYPERLINK`, `RTD`, `STOCKHISTORY`, the `CUBE` functions and `PY`, XLL add-in functions, the Excel 4.0 macro functions, functions that send cell data to an AI service, references to another workbook or a file path, and formulas the parser cannot read. Defined names are checked the same way. The error names the cell.

A host that writes trusted formulas passes `formulas: 'unrestricted'`.

Before:

```ts
await setXlsxCells(bytes, 0, [{ address: 'A1', input: '=WEBSERVICE("https://example.com")' }]);
```

After:

```ts
await setXlsxCells(bytes, 0, [{ address: 'A1', input: '=WEBSERVICE("https://example.com")' }], {
	formulas: 'unrestricted',
});
```

A server for trusted MCP callers registers the Excel tools with the option. No tool parameter turns the check off.

Before:

```ts
registerAutomationTools(server, xlsxTools, options);
```

After:

```ts
import { createXlsxTools } from 'ooxml-core/automation';

registerAutomationTools(server, createXlsxTools({ formulas: 'unrestricted' }), options);
```

### File tools refuse paths Windows would reinterpret

On Windows, the file tools refuse path segments that contain `:`, `<`, `>`, `"`, `|`, `?` or `*`, that end in a dot or a space, or that are device names such as `NUL` or `COM1`. They also refuse `/x`, which is rooted on the current drive, and `C:x`, which is relative to a drive's current folder. On every platform they refuse UNC and device paths, and paths that start with a backslash.

Before:

```json
{ "path": "C:reports\\q3.xlsx" }
```

After:

```json
{ "path": "reports\\q3.xlsx" }
```

Use a path relative to the root, or a full absolute path inside it with its drive and a separator after the colon, such as `C:\\work\\reports\\q3.xlsx`. Rename files whose names Windows would reinterpret.

## OpenTeams

This section covers `openteams-server`, the `openteams-*-viewer` packages and `ooxml-core/teams`. Deploy the app and the server together: the 2.0 client and server speak Teams protocol version 2, and a server closes clients that announce another version.

### Migrate the data directory

Rooms, file scopes and uploader records are stored under new names. The new server renames them itself when it first starts, and an older server cannot read the result.

Before:

```sh
npx openteams-server@1
```

After:

```sh
cp -r data data-backup-1.x
npx openteams-server@2
```

The migration continues on the next start if it is interrupted. After it:

- Existing chats keep working for the people their rooms list, and 1:1 chats keep their history. A chat that never got a room on the server under an old id cannot be started any more, so start a new one.
- The migration cannot tell who should be in an existing room. If you suspect someone joined a private chat without being invited, check that room's member list.
- Data stored under a chat id that only 2.0 issues cannot be genuine. It moves to a `.quarantine` folder inside `rooms`, `files` or `file-meta`, and the server logs each entry. Review or delete those folders. The server never reads them.
- Local snapshots are kept per protocol version, so changes still unsent when the app was upgraded are discarded.

### A session secret must be long

The server refuses to start when `TEAMS_SESSION_SECRET` (`sessionSecret`) is shorter than 32 characters or has fewer than 8 distinct characters. Leaving it unset still gives a random secret per process.

Before:

```sh
TEAMS_SESSION_SECRET=changeme npx openteams-server
```

After:

```sh
openssl rand -hex 32 > session-secret   # once; keep it outside the repository
TEAMS_SESSION_SECRET="$(cat session-secret)" npx openteams-server
```

### A server with only `authorizeSync` offers sync alone

A server with an `authorizeSync` hook and no `users` or `authenticate` hook now answers `POST /auth/session`, every `/files/...` request and every `/signal/...` socket with 404, because it cannot tell who a request comes from. If its clients use file storage or calls, add `authenticate`.

Before:

```js
createTeamsServer({ authorizeSync: checkTicket });
```

After:

```js
createTeamsServer({
	authorizeSync: checkTicket,
	authenticate: async ({ user, credential, headers }) =>
		(await myIdp.check(headers, credential)) ? { id: user } : null,
});
```

For development only, pass `insecureDev: true` or start with `--insecure-dev`.

### New default limits

Files now have a ceiling of 1 GiB per uploader across every scope and 20 GiB for the whole directory, which used to be off. Stored rooms have a ceiling of 10 GiB and 10,000 rooms. Presence is capped at 64 KiB per client, 4 MiB per room and 256 KiB per second per socket. A client with more than 30 costly refusals in a minute gets 429 for the rest of the minute. The largest sync message defaults to 32 MiB or a 64th of the heap, whichever is smaller. Large installations raise the limits they need.

Before:

```sh
npx openteams-server
```

After:

```sh
TEAMS_USER_QUOTA_MB=4096 \
TEAMS_FILES_QUOTA_MB=102400 \
TEAMS_ROOMS_QUOTA_MB=51200 \
TEAMS_MAX_STORED_ROOMS=50000 \
NODE_OPTIONS=--max-old-space-size=2048 \
npx openteams-server
```

A server with a small heap that hosts very large documents needs a larger heap, or `maxSyncBytes` set explicitly, to accept their first sync. With accounts, an upload to a workspace scope that names no existing room is answered with 404.

### Call and chat ids

A private chat's call uses the signaling room `call-` followed by the chat's room name (`chatCallRoomId`). The server admits only the chat's members to it and refuses the old `call-<workspace>-<chat id>` form. With accounts, a signaling socket may speak only as its signed-in user: a signal with another `from` closes the socket with `1008 peer id is not the signed-in user`.

A new 1:1 chat id is `dm-` and 32 hex digits, and a new group chat id is `g-` with a random part bound to its creator. Meeting chats are group chats. `meetingChatId` is removed, and `createGroup` no longer takes an `id`.

Before:

```ts
const chatId = meetingChatId(meeting.id);
chats.createGroup({ others, name, id: chatId });
```

After:

```ts
const chatId = meeting.chatId;
const chat = chats.createGroup({ others, name }); // the id is made for its creator
```

A third-party signaling server checks each signal's `from` against the signed-in user, and admits to a call room only the members of its chat room. `chatRoomOfCall(room)` in `ooxml-core/teams` gives the chat room of a call room.

### Servers built on `validateTeamsUpdate`

`openteams-server` does all of this already. A server of your own that uses `validateTeamsUpdate` changes in these ways:

- It must seed new chat rooms itself. The first writer can no longer fill an empty `teams:chat-members` map, and `?members=` only checks the two people of a 1:1 room against its id.
- The verdict can be `{ ok: true, reverts }`. Apply the update and then `applyUpdateReverts(doc, reverts)` in one transaction whose origin is not the sender's connection.
- A message that cannot be read or applied closes its socket with `1008 update rejected: invalid-update`.

Before:

```ts
const verdict = validateTeamsUpdate(doc, update, actor);
if (verdict.ok) Y.applyUpdate(doc, update, connection);
```

After:

```ts
const verdict = validateTeamsUpdate(doc, update, actor);
if (!verdict.ok) return socket.close(1008, 'update rejected'); // `verdict.violations` says why
doc.transact(() => {
	Y.applyUpdate(doc, update);
	applyUpdateReverts(doc, verdict.reverts ?? []);
}, serverOrigin);
```

Follow the private chat rules in [teams-area.md](teams-area.md).

`createChatRooms` passes `openRoom` a third argument, `onSynced`, and takes an optional `directorySync`. The handle `openRoom` returns must have `synced()`.

Before:

```ts
createChatRooms({ doc, base, user, openRoom: (chat, chatDoc) => connect(chat, chatDoc) });
```

After:

```ts
createChatRooms({
	doc,
	base,
	user,
	openRoom: (chat, chatDoc, onSynced) => {
		const provider = connect(chat, chatDoc);
		provider.once('sync', onSynced);
		return { destroy: () => provider.destroy(), synced: () => provider.synced };
	},
});
```

## OOXML Office deployments

This section is for hosts that deploy the OOXML Office app with their own `ai-config.json`.

### Shared chats get short-lived tokens

`collaboration.token` is ignored. Serve short-lived tokens from your own endpoint and name it in `collaboration.credentialsUrl`. The suite calls it with the user's cookies before every connection and expects `{ "token": "...", "expiresAt": "<ISO 8601 time>" }`. See [suite-ai-config.md](suite-ai-config.md#shared-chats).

Before:

```json
{
	"collaboration": {
		"syncUrl": "wss://sync.example.com",
		"token": "team-wide-secret"
	}
}
```

After:

```json
{
	"collaboration": {
		"syncUrl": "wss://sync.example.com",
		"credentialsUrl": "https://office.example.com/api/sync-token"
	}
}
```

### AI keys last for one tab

AI keys are kept for the current tab only. The **Remember the key on this device** option is gone, and keys an earlier version remembered are deleted when the app loads, so users enter them again once per tab. To spare users the key, run a gateway that authenticates them and adds the provider key on the server, and point a connection at it with `"keyMode": "none"`, as [suite-ai-config.md](suite-ai-config.md#connections) describes.

Before:

```json
{
	"connections": [{ "id": "openai", "label": "OpenAI", "provider": "openai", "keyMode": "user" }]
}
```

After:

```json
{
	"connections": [
		{
			"id": "team",
			"label": "Company models",
			"provider": "gateway",
			"endpoint": "https://ai.example.com/v1/chat/completions",
			"keyMode": "none"
		}
	]
}
```

### Shared assistant chats are display only for other people

A shared chat now carries only what every participant sees: the message text, the model's label, the names of attached files and the number of proposed edits and new files. Earlier shared conversations show their messages as display only once the room syncs, so their proposals can no longer be applied from the chat. Ask again to get proposals you can apply.

### Pages on the suite's origin need its script policy

Every page of the suite, its 404 page and each viewer's docs and demos now carry one Content Security Policy that allows script only from the site itself and never inline. A host that puts its own page on the same origin as the suite gives that page the policy from `scripts/csp.mjs`, or moves it to another origin.

Before:

```html
<script>
	initMyPage();
</script>
```

After:

```html
<meta http-equiv="Content-Security-Policy" content="<the policy from scripts/csp.mjs>" />
<script src="./my-page.js"></script>
```
