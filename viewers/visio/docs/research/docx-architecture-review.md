# DOCX architecture review for Visio

Reference: [ChristopherVR/docx-viewer](https://github.com/ChristopherVR/docx-viewer) at commit `29e25e9a32791f30ad151299389365d70af0780b`. Reviewed 2026-10-02. This is a source and targeted lifecycle review, not a claim about every browser or framework version.

## Architecture to retain

The strongest DOCX standard is one implementation with thin adapters and explicit ownership. The parser, model, layout and serialization belong to a canonical core. A browser surface then owns presentation and browser resources. React, Vue, Angular, Svelte, Solid and vanilla bindings translate lifecycle, properties, events and refs, rather than implement document logic six times.

Visio follows that ownership boundary:

- `ooxml-core/visio`: OPC/ZIP, XML, ShapeSheet cached values, inheritance, geometry, theme records, typed scenes and diagnostics.
- `ViewerController`: DOM-independent document/view state, asynchronous request identity, selection and events.
- SVG renderer: browser presentation, font measurement, raster resources and compatibility notes.
- `<visio-viewer>`: shared browser UI, keyboard behavior and resource ownership.
- Shared mount binding: one property/event contract and one mount/update/load/fit/destroy implementation.
- Native adapters: framework lifecycle and forwarding only.

This is stronger than framework neutrality alone. A shared custom element can still entangle state and rendering. Keeping the controller independently usable makes another renderer possible without recreating document state or asynchronous load semantics.

## Good reference patterns

1. **Explicit package ownership.** DOCX's instructions prohibit copying format logic into its viewer. Core entry points are re-exports rather than parallel implementations.
2. **Shared binding implementation.** Every framework passes through the same mount function. Toolbar and document commands do not live in framework templates.
3. **Exhaustive property inventories.** Compile-time completeness checks make newly added properties visible to every adapter. A mapped event type avoids separate handwritten callback lists.
4. **Initialization order.** Attach callbacks, apply initial properties, then connect the element. Initialization does not silently miss host events.
5. **Latest-load-wins behavior.** Slow work from an earlier file must not replace a newer accepted file.
6. **Injected measurement.** Browser font measurement belongs in a presentation layer, outside the canonical format model.
7. **Scoped visual tokens.** A shadow root and semantic colors separate the component from host CSS.
8. **Package-consumer tests.** Inspect tarballs, install in a clean consumer and verify exports, declarations and native entry points. Workspace success alone does not prove a usable artifact.
9. **Honest compatibility reporting.** A displayed document is not evidence of lossless import, accurate rendering or native round-trip support.

## Problems the Visio design avoids

### Stale document identity suppression

A reference binding tracks the last emitted model to suppress feedback. The sequence A → local B → external C → external B can leave C visible because B remains permanently classified as an echo.

The Visio adapters instead compare the latest native input snapshot with the previous input snapshot and forward explicit patches. An unrelated callback or toolbar rerender cannot reset an imperatively loaded document or user zoom. Reintroducing a previously omitted property requests it again. No whole-document JSON comparison is used.

Tests cover A/local-B/C/B across actual React, Vue, Angular, Svelte and Solid mounts. Models should be stored immutably; Vue shallow refs and Svelte raw state avoid creating fresh deep proxies that look like new document identities.

### Asynchronous reads outside request identity

Tracking only parser promises is insufficient. Reading a Blob is also asynchronous. A newer load, external replacement, unmount or destroy can occur before its bytes arrive.

Visio's `loadSource` gives byte reading and parsing one epoch and one loading/error state. Oversized sources are rejected before byte allocation. Old reads and failures cannot overwrite current state. Dedicated parser workers are terminated on replacement, cancellation and disposal.

### Observer exceptions masquerading as parse failures

If state subscribers and success callbacks execute inside the parser's catch block, a host callback exception can turn a successfully parsed document into a reported parse error. Reentrant subscribers can also replace the document before an old success event is emitted.

Visio isolates host listener failures, snapshots listener collections and invalidates stale notifications after a newer state revision. A subscription whose immediate callback fails is removed. Load completion checks request identity again after notifying state subscribers.

### Incomplete teardown

Destroying only the visible DOM does not dispose font listeners, observers, timers, object URLs, parser workers or retained controls. The reference also contained disposal gaps around print-layout and overflow resources.

The Visio lifecycle explicitly owns these resources. Font completion triggers text reflow while connected. Disconnect removes the listener, cancels parsing and releases image URLs. Reconnection rebuilds presentation from retained document state. Final destroy is idempotent, aborts DOM listeners and guards later commands.

### Callback errors suppressing native events or cleanup

Vue and Angular callback-map errors must not suppress their equivalent native output. Solid ref-release errors must not prevent destruction or transfer of the new handle.

Adapters use `finally` around these independent obligations. Native lifecycle tests exercise throwing callback and ref paths, StrictMode repetition, callback removal and stale work after unmount/remount.

### Save completion acknowledging newer unsaved edits

The DOCX review identified a future editor hazard: serialize revision 1, edit revision 2, then unconditionally mark the document clean when revision 1 finishes.

Visio is currently read-only. Any future writer must capture `{documentId, revision}` with an immutable snapshot. Serialization, initiating a download and confirmed host persistence must remain separate. An old save cannot mark a new revision or replacement document clean.

## Resource limits are part of renderer correctness

A bounded ZIP does not imply bounded browser work. The review led to independent scene dimensions/counts, metadata lengths, decoded-image budgets and text measurement budgets. Repeated shared images reuse one object URL. Long text uses bounded cached metrics; equal-style fragments coalesce. Paragraph run lookup uses an index rather than restarting a scan for every paragraph.

These limits protect responsiveness and make rejection explicit. They are not a security certification or proof that every valid large Visio drawing can be displayed.

## Verification boundaries

Native adapters have actual framework-mounted DOM tests, separate no-DOM import tests and selected React/Vue hydration tests. They remain private source integrations, not published standalone framework packages. Production artifact checks verify the root viewer's ESM API, TypeScript declarations, a fresh Vite consumer and its packaged parser worker.

Real Chromium interaction was blocked by this execution environment. DOM tests, source audits and static SVG rasterization do not replace browser, assistive-technology, multi-version peer or Microsoft Visio reference testing. See the capability ledger and verification report for current evidence.
