import { readFile } from 'node:fs/promises';
import { describe, expect, it } from 'vitest';
import { fieldProperties } from './doc-property-field';
import { loadDocx } from './parse';
import { removeDocumentInformation } from './remove-document-information';

// doc-property.docx: Word 16 set core, extended and custom properties and inserted DOCPROPERTY
// fields for them (`__fixtures__/computed-fields/generate-word-computed-fields.ps1`).
const bytes = await readFile(
	new URL('./__fixtures__/computed-fields/doc-property.docx', import.meta.url),
);

describe('Inspect Document and DOCPROPERTY fields', () => {
	it('stops fields showing the properties Remove All took out, through save and reload', async () => {
		const loaded = await loadDocx(bytes);
		const { model } = loaded;
		const before = fieldProperties(model);
		expect(before.app.company || before.core.creator || before.custom.length).toBeTruthy();

		const removed = removeDocumentInformation(model, ['properties']);
		const after = fieldProperties(removed);
		expect(after.core.creator).toBeUndefined();
		expect(after.core.title).toBeUndefined();
		expect(after.app.company).toBeUndefined();
		expect(after.app.manager).toBeUndefined();
		expect(after.custom).toEqual([]);
		// Word keeps these.
		expect(after.core.created).toBe(before.core.created);

		const reloaded = fieldProperties((await loadDocx(await loaded.save(removed))).model);
		expect(reloaded.core.creator ?? '').toBe('');
		expect(reloaded.app.company ?? '').toBe('');
		expect(reloaded.custom).toEqual([]);
	});
});
