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
- `meta.reason` is `'save'` for a `.pptx` save and `'saveAs'` for any other format; the viewer's
  Save button and backstage do not otherwise tell the two apart.

A rejected `save()` is treated as a failed save: the AutoRecover copy is kept (it is only discarded
after a save succeeds and Options > Save asks for it), and the error is logged.

The bytes are the same ones `getContent()` returns, so a password-protected deck is handed over
encrypted.

## Recovery copies

Recovery autosave (the `autosave` and `autosaveIntervalMs` props, keyed by `filePath`) normally
writes to the library's own IndexedDB database. When `storage` provides `readRecovery`,
`writeRecovery` and `clearRecovery`, every recovery path uses them instead: the autosave writer,
the recovery prompt shown on load (and its Discard), the cleanup after a successful save, Version
History, and `restoreSessionDeck(storage)`. The entry id is the `filePath` key; `savedAt` is a
millisecond timestamp. A storage that provides only some of the three methods is ignored for
recovery, and IndexedDB stays in use.

Not routed through `storage`: File > Open > Recent and Account > Storage & Privacy describe the
browser's own local data and keep reading IndexedDB, and the expiry sweep of old IndexedDB
snapshots still runs on that database only (a host keeps its own retention).

## No `file-command` event

The Word, Excel and Visio elements announce a cancelable `file-command` event before Save, Open
and New. The PowerPoint bindings have no such event: to take over saving, pass `storage`; to take
over File > Open, use `onOpenFile` (Vue and Angular: the `onOpenFile` input).
