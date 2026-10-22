# Host-owned storage

Every binding takes a `storage` prop (Angular: input, Vanilla: option) that hands the viewer's file
I/O to the host. It is the `DocumentStorage` contract from `ooxml-core/embedding`, the same one the
Word, Excel and Visio editors accept, and is re-exported as a type from `ooxml-ui/pptx`.

```ts
import type { DocumentStorage } from 'ooxml-ui/pptx';

const storage: DocumentStorage = {
	// Required: File > Save and Save As.
	async save(bytes, { fileName, mimeType, reason }) {
		await uploadToMyServer(fileName, bytes, mimeType);
		// Optional: report the name the document now has.
		return { fileName };
	},
	// Optional, all three together: AutoRecover copies.
	readRecovery: (id) => myRecovery.get(id),
	writeRecovery: (entry) => myRecovery.put(entry),
	clearRecovery: (id) => myRecovery.delete(id),
	// Optional: lets Version History list the kept copies.
	listRecovery: () => myRecovery.list(),
};
```

## Save and Save As

Without `storage` the viewer downloads the saved file. With it, Save and every Save As format
(`.pptx`, `.ppsx`, `.pptm`, `.ppt`) call `storage.save(bytes, meta)` instead:

- `meta.fileName` is the source stem plus the extension of the format written
  (`savedPresentationFileName`), so `report.ppt` saves as `report.pptx`.
- `meta.mimeType` follows the format (`.ppt` is the OLE2 type, `.ppsx` and `.pptm` have their own).
- `meta.reason` is the command that ran: `'save'` for File > Save (the title bar, the backstage
  Save item, the mobile toolbar and Vanilla/Svelte `downloadPptx()`), `'saveAs'` for every Save As
  entry, including Save As > PowerPoint Presentation (.pptx).

A failed save (the host's `save()` rejected, or serialising failed) is reported the same way in
all five bindings: the binding's `error` event (`onError`, `@error`, `(error)`, `onerror`) receives
`Save failed: <reason>`, the user sees the "Save failed" notice in the viewer's notice stack, and
the AutoRecover copy is kept (it is only discarded after a save succeeds and Options > Save asks
for it). The save methods resolve rather than reject.

The bytes are the same ones `getContent()` returns, so a password-protected deck is handed over
encrypted.

## Recovery copies

Recovery autosave (the `autosave` and `autosaveIntervalMs` props, keyed by `filePath`) normally
writes to the library's own IndexedDB database. When `storage` provides `readRecovery`,
`writeRecovery` and `clearRecovery`, every recovery path uses them instead: the autosave writer,
the recovery prompt shown on load (and its Discard), the cleanup after a successful save, Version
History, File > Open > Recent, Account > Storage & Privacy and `restoreSessionDeck(storage)`. The
entry id is the `filePath` key; `savedAt` is a millisecond timestamp. A storage that provides only
some of the three methods is ignored for recovery, and IndexedDB stays in use.

Listing (Recent, Account usage, Version History) needs `listRecovery`; without it those lists are
empty. The housekeeping of Options > Save ("Delete cached files", "clear cache on close" and the
"cache retention" sweep) clears the host's copies only when it provides `listRecovery`. Without it
the viewer leaves the host's copies alone, and retention is the host's job.

## No `file-command` event

The Word, Excel and Visio elements announce a cancelable `file-command` event before Save, Open
and New. The PowerPoint bindings have no such event: to take over saving, pass `storage`; to take
over File > Open, use `onOpenFile` (Vue and Angular: the `onOpenFile` input).
