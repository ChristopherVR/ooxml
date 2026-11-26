import { describe, expect, it } from 'vitest';
import { documentParagraphs } from '../cell-blocks';
import { plainText } from '../document-stats';
import { createDocument, type DocumentModel, type Table, type TextRun } from '../model';
import { loadDocx } from '../parse';
import { saveDocx } from '../save';
import {
	finishMailMerge,
	finishMailMergeIndividually,
	mailMergeAddressBlock,
	loadMailMergeStationery,
	mailMergeField,
	mailMergeGreetingLine,
	mailMergeRule,
	previewMailMerge,
	resolveMailMergePrompts,
} from './index';
import { complex, code, paragraph, template } from './test-fixture';

const { addEnvelope, createLabels, updateLabels } = await loadMailMergeStationery();

const records = {
	records: [
		{
			Given: 'Ada',
			LastName: 'Lovelace',
			City: 'Paris',
			PostalCode: '75001',
			CountryorRegion: 'France',
		},
		{
			Given: 'Bea',
			LastName: '',
			City: 'Boston',
			State: 'MA',
			PostalCode: '02110',
			CountryorRegion: 'USA',
		},
	],
};

const documentWith = (runs: TextRun[]): DocumentModel => ({
	...createDocument(),
	blocks: [{ type: 'paragraph', id: 'p1', runs }],
});

describe('mail merge stationery', () => {
	it('prepends an envelope and round-trips its dimensions, positions, tray and letter', async () => {
		const letter = documentWith([{ text: 'The original letter' }]);
		const changed = addEnvelope(letter, {
			size: '10',
			delivery: 'Ada\nLondon',
			returnAddress: 'Sender',
			firstPageTray: 15,
			deliveryPosition: { leftTwips: 5600, topTwips: 2850 },
		});
		const { model } = await loadDocx(await saveDocx(changed));
		expect(model.sections).toMatchObject([
			{
				pageWidthTwips: 13680,
				pageHeightTwips: 5940,
				orientation: 'landscape',
				paperSource: { first: 15 },
				pageNumbering: { start: 0 },
			},
			{ pageWidthTwips: 12240, pageHeightTwips: 15840 },
		]);
		expect(model.blocks[1]).toMatchObject({ frame: { xTwips: 5600, yTwips: 2850 } });
		expect(plainText(model)).toContain('Sender\nAda\nLondon');
		expect(plainText(model)).toContain('The original letter');
		const envelope = addEnvelope(
			letter,
			{ size: 'dl', delivery: '', returnAddress: 'Hidden', omitReturnAddress: true },
			'merge',
		);
		const { model: reopened } = await loadDocx(await saveDocx(envelope));
		expect(reopened.sections).toHaveLength(1);
		expect(reopened.mailMerge?.mainDocumentType).toBe('envelopes');
		expect(plainText(reopened)).toBe('');
	});

	it('writes a fixed label grid with gutters, prints one position or every label, and updates templates', async () => {
		const full = await loadDocx(await saveDocx(createLabels({ product: '5160', text: 'Same' })));
		expect(plainText(full.model).match(/Same/g) ?? []).toHaveLength(30);
		const single = await loadDocx(
			await saveDocx(
				createLabels({ product: '5160', text: 'Only', single: { row: 2, column: 3 } }),
			),
		);
		expect(single.model.blocks[0]).toMatchObject({
			type: 'table',
			layout: 'fixed',
			grid: [3780, 180, 3780, 180, 3780],
		});
		const table = single.model.blocks[0];
		if (table?.type !== 'table') throw new Error('Missing label table');
		expect(table.rows[1]?.[4]?.paragraphs[0]?.runs[0]?.text).toBe('Only');
		expect(plainText(single.model).match(/Only/g) ?? []).toHaveLength(1);
		expect(table.rowProperties?.[0]).toMatchObject({ heightTwips: 1440, heightRule: 'exact' });

		const blanks = createLabels({ product: '5163' });
		const first = documentParagraphs(blanks.blocks)[0];
		if (!first) throw new Error('Missing first label');
		first.runs = [{ text: 'To: ' }, mailMergeField('Given')];
		const updated = updateLabels({ ...blanks, mailMerge: { mainDocumentType: 'mailingLabels' } });
		const loaded = await loadDocx(await saveDocx(updated));
		const paragraphs = documentParagraphs(loaded.model.blocks);
		expect(paragraphs[1]?.runs.map(({ text }) => text).join('')).toBe('');
		expect(paragraphs[2]?.runs.some(({ fieldCode }) => fieldCode?.trim() === 'NEXT')).toBe(true);
		expect(new Set(paragraphs.map(({ id }) => id)).size).toBe(paragraphs.length);
		const result = finishMailMerge(loaded.model, records);
		expect(result).toMatchObject({ mergedCount: 1, issues: [] });
		const reopened = await loadDocx(await loaded.save(result.document));
		expect(plainText(reopened.model)).toContain('To: Ada\t\tTo: Bea');
		expect(plainText(reopened.model).match(/To:/g) ?? []).toHaveLength(2);
		expect((reopened.model.blocks[0] as Table).rows).toHaveLength(5);
	});
});

describe('native mail merge fields and prompts', () => {
	it('round-trips native fields and evaluates mappings, country and greeting switches', async () => {
		const source = documentWith([
			...mailMergeAddressBlock({ country: '2', excludeCountry: 'France', formatByCountry: true }),
			{ text: '\n\n' },
			...mailMergeGreetingLine({
				name: 'full',
				before: 'Hello',
				after: '!',
				fallback: 'Hello there!',
			}),
		]);
		source.mailMerge = { fieldMappings: [{ mappedName: 'FirstName', name: 'Given' }] };
		const loaded = await loadDocx(await saveDocx(source));
		const fields = documentParagraphs(loaded.model.blocks)[0]?.runs.flatMap(
			(run) => run.field?.instr ?? [],
		);
		expect(fields).toEqual([
			expect.stringContaining('ADDRESSBLOCK'),
			expect.stringContaining('GREETINGLINE'),
		]);
		expect(plainText(previewMailMerge(loaded.model, records, 1).document)).toBe(
			'Ada Lovelace\n75001 Paris\n\nHello Ada Lovelace!',
		);
		expect(plainText(previewMailMerge(loaded.model, records, 2).document)).toBe(
			'Bea\nBoston, MA 02110\nUSA\n\nHello Bea!',
		);
	});

	it('uses one ASK answer, asks FILLIN per record, supports cancellation and keeps defaults during preview', async () => {
		const source = documentWith([
			...mailMergeRule({
				kind: 'ASK',
				bookmark: 'Who',
				prompt: 'Who?',
				answer: 'Default',
				once: true,
			}),
			...mailMergeRule({ kind: 'REF', bookmark: 'Who' }),
			{ text: ' ' },
			...mailMergeRule({ kind: 'FILLIN', bookmark: '', prompt: 'Where?', answer: 'Here' }),
		]);
		const loaded = await loadDocx(await saveDocx(source));
		expect(plainText(previewMailMerge(loaded.model, records, 1).document)).toBe('Default Here');
		const seen: string[] = [];
		const options = await resolveMailMergePrompts(loaded.model, records, async (request) => {
			seen.push(`${request.kind}:${request.recordNumber}`);
			return request.kind === 'ASK' ? 'Answer' : String(request.recordNumber);
		});
		expect(seen).toEqual(['ASK:1', 'FILLIN:1', 'FILLIN:2']);
		expect(
			finishMailMergeIndividually(loaded.model, records, options).documents.map(({ document }) =>
				plainText(document),
			),
		).toEqual(['Answer 1', 'Answer 2']);
		expect(
			await resolveMailMergePrompts(loaded.model, records, async () => undefined),
		).toBeUndefined();
		let calls = 0;
		finishMailMerge(loaded.model, records, {
			answerPrompt: ({ once }) => {
				if (once) calls++;
				return '';
			},
		});
		expect(calls).toBe(1);
	});

	it('prompts only for the IF branch chosen by a live answer', async () => {
		const fill = (label: string) => complex(code(` FILLIN "${label}" \\d "no" `), 'no');
		const loaded = await template(
			paragraph(
				complex(
					code(' IF "') +
						fill('Branch?') +
						code('" = "yes" "') +
						fill('Yes?') +
						code('" "') +
						fill('No?') +
						code('" '),
					'cached',
				),
			),
		);
		const seen: string[] = [];
		const options = await resolveMailMergePrompts(
			loaded.model,
			{ records: [{}] },
			async ({ text }) => {
				seen.push(text);
				return text === 'Branch?' ? 'yes' : 'chosen';
			},
		);
		expect(seen).toEqual(['Branch?', 'Yes?']);
		expect(
			finishMailMergeIndividually(loaded.model, { records: [{}] }, options).documents.map(
				({ document }) => plainText(document),
			),
		).toEqual(['chosen']);
	});
});
