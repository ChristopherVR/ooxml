import { readFile } from 'node:fs/promises';
import { describe, expect, it } from 'vitest';
import { computedFields } from './computed-fields';
import { docPropertyResult, NO_PROPERTY_NAME, UNKNOWN_PROPERTY } from './doc-property-field';
import { fieldResultSpans } from './field-result-spans';
import type { DocumentModel, Paragraph } from './model';
import { loadDocx } from './parse';

// doc-property.docx: Word 16 inserted and updated each DOCPROPERTY field
// (`__fixtures__/computed-fields/generate-word-computed-fields.ps1`), so every cached result
// is Word's own.
const bytes = await readFile(
	new URL('./__fixtures__/computed-fields/doc-property.docx', import.meta.url),
);

/** Each field of the body: its instruction, Word's cached result and the computed one. */
function fieldsOf(model: DocumentModel, errors = true) {
	const fields = computedFields(model, { errors });
	return model.blocks.flatMap((block) =>
		block.type === 'paragraph'
			? fieldResultSpans(block.runs).map((span) => ({
					instr: span.instr,
					word: span.runs.map((at) => block.runs[at]!.text).join(''),
					computed: fields.result(span, block)?.text,
				}))
			: [],
	);
}

/** Results Word draws from what it counts or from the system's regional date format. */
const KEPT = new Set(['DOCPROPERTY Due', 'DOCPROPERTY Pages']);

describe('DOCPROPERTY', () => {
	it("computes Word's results for core, extended and custom properties", async () => {
		const { model } = await loadDocx(bytes);
		const fields = fieldsOf(model);
		expect(fields.length).toBeGreaterThan(50);
		for (const field of fields)
			if (KEPT.has(field.instr)) expect(field.computed, field.instr).toBeUndefined();
			else expect(field.computed, field.instr).toBe(field.word);
	});

	it('keeps the result of a field naming no property outside Update Fields', async () => {
		const { model } = await loadDocx(bytes);
		const missing = fieldsOf(model, false).filter((field) => field.word.startsWith('Error!'));
		expect(missing.length).toBeGreaterThan(1);
		for (const field of missing) expect(field.computed).toBeUndefined();
	});

	it('reports a missing or unknown property as Word does', () => {
		const set = { core: {}, app: {}, custom: [] };
		expect(docPropertyResult('DOCPROPERTY', set)).toBe(NO_PROPERTY_NAME);
		expect(docPropertyResult('DOCPROPERTY Nope', set)).toBe(UNKNOWN_PROPERTY);
		expect(docPropertyResult('DOCPROPERTY Title', set)).toBe('');
		expect(docPropertyResult('PAGE', set)).toBeUndefined();
	});

	it('keeps the result of a field with a switch it does not evaluate', () => {
		const set = { core: {}, app: {}, custom: [{ name: 'n', type: 'i4' as const, value: 7 }] };
		expect(docPropertyResult('DOCPROPERTY n \\# "00"', set)).toBeUndefined();
		expect(docPropertyResult('DOCPROPERTY n \\* Bogus', set)).toBeUndefined();
		expect(docPropertyResult('DOCPROPERTY n \\* OrdText \\* Upper', set)).toBe('SEVENTH');
	});

	it('shows the core properties as edited in the document', async () => {
		const { model } = await loadDocx(bytes);
		const edited = { ...model, properties: { ...model.properties, title: 'Edited' } };
		const title = fieldsOf(edited).find((field) => field.instr === 'DOCPROPERTY Title');
		expect(title?.computed).toBe('Edited');
		expect(
			fieldsOf({ ...edited, properties: {} }).find((field) => field.instr === 'DOCPROPERTY Title')
				?.computed,
		).toBe('');
	});

	it('writes the results back on save, keeping locked fields', async () => {
		const loaded = await loadDocx(bytes);
		const { model } = loaded;
		const title = (block: DocumentModel['blocks'][number]) =>
			block.type === 'paragraph' &&
			block.runs.some((run) => run.field?.instr.trim() === 'DOCPROPERTY Title');
		const locked = model.blocks.findIndex(title);
		const blocks = model.blocks.map((block, index) => {
			if (index !== locked) return block;
			const paragraph = block as Paragraph;
			return {
				...paragraph,
				runs: paragraph.runs.map((run) =>
					run.field ? { ...run, fieldFlags: { locked: true } } : run,
				),
			};
		});
		const saved = await loaded.save({
			...model,
			blocks,
			properties: { ...model.properties, title: 'Saved title', creator: 'Grace Hopper' },
		});
		const reloaded = fieldsOf((await loadDocx(saved)).model);
		const result = (instr: string) => reloaded.filter((field) => field.instr === instr);
		expect(result('DOCPROPERTY Title').map((field) => field.word)).toEqual(['my report title']);
		expect(result('DOCPROPERTY title')[0]?.word).toBe('Saved title');
		expect(result('DOCPROPERTY Author')[0]?.word).toBe('Grace Hopper');
		expect(result('DOCPROPERTY  Author  \\* Upper \\* MERGEFORMAT')[0]?.word).toBe('GRACE HOPPER');
		expect(result('DOCPROPERTY Company')[0]?.word).toBe('Acme Ltd');
	});
});
