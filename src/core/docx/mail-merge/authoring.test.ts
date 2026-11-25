import { describe, expect, it } from 'vitest';
import { documentParagraphs } from '../cell-blocks';
import { createDocument, type TextRun } from '../model';
import { loadDocx } from '../parse';
import { saveDocx } from '../save';
import {
	checkMailMerge,
	composeMailMergeFields,
	finishMailMerge,
	mailMergeDelimitedText,
	mailMergeField,
	mailMergeRule,
	mailMergeWithinWorkLimit,
	previewMailMerge,
	setMailMergeFieldMapping,
	type MailMergeFieldMapping,
} from './index';

function model(runs: TextRun[]) {
	return { ...createDocument(), blocks: [{ type: 'paragraph' as const, id: 'p1', runs }] };
}

describe('authored merge templates', () => {
	it('round-trips quoted field names and nested IF operands with literal quotes and slashes', async () => {
		const key = 'Status "current" \\';
		const source = model([
			mailMergeField(key),
			{ text: ': ' },
			...mailMergeRule({
				kind: 'IF',
				field: key,
				operator: '=',
				value: 'yes',
				yes: 'Use "A" \\ B',
				no: 'Skip',
			}),
		]);
		const reopened = await loadDocx(await saveDocx(source));
		const data = { records: [{ [key]: 'yes' }, { [key]: 'no' }] };
		expect(checkMailMerge(reopened.model, data)).toEqual([]);
		const render = (n: number) =>
			documentParagraphs(previewMailMerge(reopened.model, data, n).document.blocks)[0]
				?.runs.map((run) => run.text)
				.join('');
		expect(render(1)).toBe('yes: Use "A" \\ B');
		expect(render(2)).toBe('no: Skip');
	});

	it('authors named prompts, defaults, record counters, and conditional record rules', async () => {
		const source = model([
			...mailMergeRule({
				kind: 'SKIPIF',
				field: 'Skip',
				operator: '=',
				value: 'yes',
				yes: '',
				no: '',
			}),
			...mailMergeRule({ kind: 'ASK', bookmark: 'Name', prompt: 'Name?', answer: 'Ada' }),
			...mailMergeRule({ kind: 'REF', bookmark: 'Name' }),
			...mailMergeRule({ kind: 'FILLIN', bookmark: '', prompt: 'Place?', answer: 'London' }),
			...mailMergeRule({ kind: 'MERGEREC' }),
			...mailMergeRule({ kind: 'MERGESEQ' }),
			...mailMergeRule({
				kind: 'NEXTIF',
				field: 'Skip',
				operator: '=',
				value: 'no',
				yes: '',
				no: '',
			}),
			mailMergeField('Name'),
			...mailMergeRule({ kind: 'NEXT' }),
		]);
		const { model: reopened } = await loadDocx(await saveDocx(source));
		const data = {
			records: [
				{ Skip: 'yes', Name: 'A' },
				{ Skip: 'no', Name: 'B' },
				{ Skip: 'no', Name: 'C' },
			],
		};
		expect(previewMailMerge(reopened, data, 1).skipped).toBe(true);
		expect(finishMailMerge(reopened, data)).toMatchObject({
			issues: [],
			mergedCount: 1,
			document: {
				blocks: [
					{
						runs: [
							{ text: 'Ada' },
							{ text: 'London' },
							{ text: '2' },
							{ text: '1' },
							{ text: 'C' },
						],
					},
				],
			},
		});
	});

	it('preserves configured mappings and continuous directory sections through Word files', async () => {
		const fields = ['Given', 'City'];
		const mappings: MailMergeFieldMapping[] = [];
		setMailMergeFieldMapping(mappings, 'FirstName', fields, 'Given');
		const source = {
			...model(composeMailMergeFields('address', fields, mappings)),
			mailMerge: { mainDocumentType: 'catalog', fieldMappings: mappings },
		};
		const loaded = await loadDocx(await saveDocx(source));
		const result = finishMailMerge(loaded.model, {
			fields,
			records: [
				{ Given: 'Ada', City: 'London' },
				{ Given: 'Bea', City: 'Paris' },
			],
		});
		expect(result.issues).toEqual([]);
		const reopened = await loadDocx(await loaded.save(result.document));
		expect(reopened.model.sections?.[1]?.type).toBe('continuous');
		expect(reopened.model.mailMerge).toBeUndefined();
		expect(
			documentParagraphs(reopened.model.blocks)[0]
				?.runs.map(({ text }) => text)
				.join(''),
		).toBe('Ada\nLondon');
	});

	it('bounds both source dimensions and repeated recipient text before creating output', () => {
		const limit = { maxRecords: 1, maxFields: 2 };
		expect(() => mailMergeDelimitedText('A,B,C\n1,2,3', { dataLimits: limit }).read()).toThrow(
			RangeError,
		);
		expect(() => mailMergeDelimitedText('A\n1\n2', { dataLimits: limit }).read()).toThrow(
			RangeError,
		);
		expect(mailMergeWithinWorkLimit(model([mailMergeField('A')]), { records: [{ A: 'ok' }] })).toBe(
			true,
		);
		expect(
			mailMergeWithinWorkLimit(model(Array.from({ length: 100 }, () => mailMergeField('A'))), {
				records: [{ A: 'x'.repeat(24_000) }],
			}),
		).toBe(false);
	});
});
