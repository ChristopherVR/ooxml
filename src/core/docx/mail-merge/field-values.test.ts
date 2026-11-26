import { describe, expect, it } from 'vitest';
import { createDocument, type TextRun } from '../model';
import { checkMailMerge, finishMailMergeIndividually, previewMailMerge } from './index';
import { code, complex, marker, paragraph, run, simple, template } from './test-fixture';

function preview(instr: string, value: string | number | Date | null, fields?: string[]) {
	const model = createDocument();
	model.blocks = [
		{
			type: 'paragraph',
			id: 'p1',
			runs: [{ text: 'saved', field: { instr, simple: true } }],
		},
	];
	const result = previewMailMerge(
		model,
		{ records: [{ Value: value }], ...(fields && { fields }) },
		1,
		{
			fieldMappings: [{ mappedName: 'Mapped', name: 'Value' }],
		},
	);
	return { text: plain(result.document.blocks[0]), issues: result.issues };
}

function plain(block: { runs: TextRun[] } | { type: 'table' } | undefined) {
	return block && 'runs' in block ? block.runs.map((run) => run.text).join('') : undefined;
}

describe('Word mail merge field evaluation', () => {
	it.each([
		['MERGEFIELD Value', '00123', '00123'],
		['MERGEFIELD Value \\b "Dear " \\f "," \\* Caps', 'ada lovelace', 'Dear Ada Lovelace,'],
		['MERGEFIELD Value \\b "Dear " \\f ","', null, ''],
		['MERGEFIELD Value \\b "No. "', 0, 'No. 0'],
		['MERGEFIELD Mapped \\m \\* Upper', 'Ada', 'ADA'],
		['MERGEFIELD Value \\v', 'A12-3', 'Ａ一二︱三'],
		['MERGEFIELD Value \\@ "MMMM d, yyyy"', '2026-11-26', 'November 26, 2026'],
		[
			'MERGEFIELD Value \\@ "yyyy-MM-dd HH:mm"',
			new Date('2026-11-26T02:30:00Z'),
			'2026-11-26 02:30',
		],
		['MERGEFIELD Value \\# "$#,##0.00"', 1234.5, '$1,234.50'],
		['MERGEFIELD Value \\# "$,0.00"', 1234.5, '$1,234.50'],
		['MERGEFIELD Value \\# $###', 15, '$ 15'],
		['MERGEFIELD Value \\# ##%', 33, '33%'],
		['MERGEFIELD Value \\# x##', 222492, '492'],
		['MERGEFIELD Value \\# .x', 0.75, '.8'],
		['MERGEFIELD Value \\# +##', -10, '-10'],
		['MERGEFIELD Value \\# "0.00;(0.00);\'zero\'"', -2.5, '(2.50)'],
		['MERGEFIELD Value \\# "0.00;(0.00);\'zero\'"', 0, 'zero'],
		['MERGEFIELD Value \\* Roman', 9, 'IX'],
		['MERGEFIELD Value \\# "0.00"', 'word', 'word'],
		['FILLIN "Your name?" \\d "Jane Doe" \\o', null, 'Jane Doe'],
		['FILLIN "Your name?" \\d ""', null, ''],
		['FILLIN "Your name?"', null, 'saved'],
	] as const)('%s merges %s', (instruction, value, expected) => {
		expect(preview(instruction, value)).toEqual({ text: expected, issues: [] });
	});

	it('reports missing fields, preserves invalid and unsupported fields, and accepts schema-defined blanks', () => {
		expect(preview('MERGEFIELD Missing', null)).toEqual({
			text: '',
			issues: [{ code: 'unknown-field', recordNumber: 1, blockId: 'p1', field: 'Missing' }],
		});
		expect(preview('MERGEFIELD Missing', null, ['Value', 'Missing']).issues).toEqual([]);
		expect(preview('MERGEFIELD Value \\b', null)).toMatchObject({
			text: 'saved',
			issues: [{ code: 'invalid-instruction' }],
		});
		expect(preview('MERGEBARCODE', null)).toMatchObject({
			text: 'saved',
			issues: [{ code: 'unsupported-field' }],
		});
		expect(checkMailMerge(createDocument(), { records: [] })).toEqual([{ code: 'no-recipients' }]);
	});

	it('uses ASK defaults in following REF fields and only evaluates the chosen nested IF branch', async () => {
		const ask = complex(code(' ASK Month "Which month?" \\d August \\o '));
		const greeting = complex(
			code(' IF "') +
				complex(code(' MERGEFIELD Name '), 'saved name') +
				code('" = "A*" "Hello ') +
				complex(code(' MERGEFIELD Name '), 'saved name') +
				code('" "') +
				complex(code(' MERGEFIELD Missing '), 'saved unused') +
				code('" '),
			'stale IF result',
		);
		const loaded = await template(paragraph(ask + simple('REF Month') + greeting));
		const data = { records: [{ Name: 'Ada "quoted" \\ literal' }] };
		const result = previewMailMerge(loaded.model, data, 1);
		expect(result.issues).toEqual([]);
		expect(plain(result.document.blocks[0])).toBe('AugustHello Ada "quoted" \\ literal');
		expect(plain(loaded.model.blocks[0])).toContain('stale IF result');
	});

	it('consumes NEXTIF records, skips matching SKIPIF output and distinguishes record and sequence', async () => {
		const field = (name: string) => complex(code(` MERGEFIELD ${name} `), 'stale');
		const rule = (name: string, key: string, test: string) =>
			complex(code(` ${name} `) + field(key) + code(` = "${test}" `));
		const loaded = await template(
			paragraph(
				rule('SKIPIF', 'Skip', 'yes') +
					field('Name') +
					rule('NEXTIF', 'Join', 'yes') +
					field('Name') +
					simple('MERGEREC') +
					simple('MERGESEQ'),
			),
		);
		const data = {
			records: [
				{ Name: 'Ada', Skip: 'yes', Join: 'no' },
				{ Name: 'Bea', Skip: 'no', Join: 'yes' },
				{ Name: 'Cal', Skip: 'no', Join: 'no' },
				{ Name: 'Dee', Skip: 'no', Join: 'no' },
			],
		};
		const result = finishMailMergeIndividually(loaded.model, data);
		expect(result.issues).toEqual([]);
		expect(
			result.documents.map(({ recordNumber, document }) => [
				recordNumber,
				plain(document.blocks[0]),
			]),
		).toEqual([
			[2, 'BeaCal31'],
			[4, 'DeeDee42'],
		]);
		expect(previewMailMerge(loaded.model, data, 1).skipped).toBe(true);
		expect(plain(previewMailMerge(loaded.model, data, 4).document.blocks[0])).toBe('DeeDee42');
	});

	it('reports a merge field spanning paragraphs and preserves its saved result', async () => {
		const longField =
			paragraph(marker('begin') + code(' MERGEFIELD Name ') + marker('separate')) +
			paragraph(run('Saved cross-paragraph value') + marker('end') + simple('MERGEFIELD Name'));
		const loaded = await template(longField);
		const preview = previewMailMerge(loaded.model, { records: [{ Name: 'Ada' }] }, 1);
		expect(preview.issues).toEqual([
			{ code: 'invalid-instruction', recordNumber: 1, blockId: 'p0', field: 'MERGEFIELD' },
		]);
		expect(plain(preview.document.blocks[1])).toBe('Saved cross-paragraph valueAda');
		expect(preview.document.blocks[0]).toBe(loaded.model.blocks[0]);
	});
});
