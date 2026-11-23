import { describe, expect, it } from 'vitest';
import { computedFields, refreshComputedFields } from './computed-fields';
import { docInfoFieldResult } from './doc-property-field';
import { ifFieldResult } from './conditional-field';
import { ZERO_DIVIDE_ERROR, formulaFieldResult } from './formula-field';
import type { DocumentModel, Paragraph } from './model';

const para = (id: string, instr: string, cached = 'stale'): Paragraph => ({
	type: 'paragraph',
	id,
	runs: [
		{ text: '', fieldChar: 'begin' },
		{ text: '', fieldCode: ` ${instr} ` },
		{ text: '', fieldChar: 'separate' },
		{ text: cached },
		{ text: '', fieldChar: 'end' },
	],
});

const emptyModel = (
	blocks: DocumentModel['blocks'],
	properties?: DocumentModel['properties'],
): DocumentModel => ({
	blocks,
	page: {
		width: 816,
		height: 1056,
		marginTop: 96,
		marginRight: 96,
		marginBottom: 96,
		marginLeft: 96,
	},
	warnings: [],
	...(properties ? { properties } : {}),
});

describe('Word document-information fields (AUTHOR, TITLE, SUBJECT, etc.)', () => {
	const props = {
		core: {
			creator: 'Ada Lovelace',
			title: 'Analytical Engine',
			subject: 'Computing',
			keywords: 'math, engine',
			description: 'Draft notes',
			lastModifiedBy: 'Charles Babbage',
			revision: '4',
			created: '2026-05-01T10:30:00Z',
			modified: '2026-06-15T12:00:00Z',
			lastPrinted: '2026-06-16T08:00:00Z',
		},
		app: {},
		custom: [],
	};

	it('resolves built-in document-information fields and applies formatting switches', () => {
		expect(docInfoFieldResult('AUTHOR', props)).toBe('Ada Lovelace');
		expect(docInfoFieldResult('AUTHOR \\* UPPER', props)).toBe('ADA LOVELACE');
		expect(docInfoFieldResult('AUTHOR "Grace Hopper"', props)).toBe('Grace Hopper');
		expect(docInfoFieldResult('TITLE', props)).toBe('Analytical Engine');
		expect(docInfoFieldResult('SUBJECT', props)).toBe('Computing');
		expect(docInfoFieldResult('KEYWORDS', props)).toBe('math, engine');
		expect(docInfoFieldResult('COMMENTS', props)).toBe('Draft notes');
		expect(docInfoFieldResult('LASTSAVEDBY', props)).toBe('Charles Babbage');
		expect(docInfoFieldResult('REVNUM', props)).toBe('4');
		expect(docInfoFieldResult('CREATEDATE \\@ "yyyy-MM-dd"', props)).toBe('2026-05-01');
		expect(docInfoFieldResult('SAVEDATE \\@ "yyyy-MM-dd"', props)).toBe('2026-06-15');
		expect(docInfoFieldResult('PRINTDATE \\@ "yyyy-MM-dd"', props)).toBe('2026-06-16');
		// Date fields without \@ keep their saved result.
		expect(docInfoFieldResult('CREATEDATE', props)).toBeUndefined();
	});
});

describe('Word IF field', () => {
	it('evaluates numeric, string, and wildcard comparisons', () => {
		expect(ifFieldResult('IF 10 > 2 "yes" "no"')).toBe('yes');
		expect(ifFieldResult('IF 2 > 10 "yes" "no"')).toBe('no');
		expect(ifFieldResult('IF 5 = 5 "equal"')).toBe('equal');
		expect(ifFieldResult('IF 5 <> 5 "diff"')).toBe('');
		expect(ifFieldResult('IF "apple" = "app*" "match" "miss"')).toBe('match');
		expect(ifFieldResult('IF "cat" = "c?t" "ok" "no"')).toBe('ok');
		expect(ifFieldResult('IF "cat" <> "dog" "different" "same" \\* UPPER')).toBe('DIFFERENT');
	});

	it('returns undefined for malformed or unsupported IF instructions', () => {
		expect(ifFieldResult('IF 1 == 1 "a" "b"')).toBeUndefined();
		expect(ifFieldResult('IF 1 = 1')).toBeUndefined();
		expect(ifFieldResult('IF 1 = 1 "1" "0" \\# "0.00"')).toBeUndefined();
	});
});

describe('Word = formula field', () => {
	it('evaluates arithmetic, precedence, comparisons, functions, and \\* formats', () => {
		expect(formulaFieldResult('= 2 + 3 * 4')).toBe('14');
		expect(formulaFieldResult('=(2+3)*4')).toBe('20');
		expect(formulaFieldResult('= -5 + 2^3')).toBe('3');
		expect(formulaFieldResult('= 50% * 80')).toBe('40');
		expect(formulaFieldResult('= SUM(10, 20, 30)')).toBe('60');
		expect(formulaFieldResult('= AVERAGE(2, 4, 6)')).toBe('4');
		expect(formulaFieldResult('= ROUND(3.14159, 2)')).toBe('3.14');
		expect(formulaFieldResult('= IF(3 > 2, MAX(1, 9), MIN(1, 9))')).toBe('9');
		expect(formulaFieldResult('= MOD(10, 3)')).toBe('1');
		expect(formulaFieldResult('= AND(1, NOT(0))')).toBe('1');
		expect(formulaFieldResult('= 42 \\* ROMAN')).toBe('XLII');
	});

	it('reports zero divide and preserves cached results for table references or \\# pictures', () => {
		expect(formulaFieldResult('= 10 / 0')).toBe(ZERO_DIVIDE_ERROR);
		expect(formulaFieldResult('= SUM(ABOVE)')).toBeUndefined();
		expect(formulaFieldResult('= A1 + B1')).toBeUndefined();
		expect(formulaFieldResult('= 10 / 4 \\# "0.00"')).toBeUndefined();
	});
});

describe('computedFields and refreshComputedFields integration', () => {
	it('evaluates AUTHOR, IF, and = in body and keeps errors suppressed unless requested', () => {
		const pAuthor = para('p1', 'AUTHOR');
		const pIf = para('p2', 'IF 3 >= 3 "pass" "fail"');
		const pEq = para('p3', '=6*7');
		const pDiv = para('p4', '=1/0', 'old');
		const model = emptyModel([pAuthor, pIf, pEq, pDiv], { creator: 'Katherine Johnson' });

		const fields = computedFields(model);
		expect(fields.result({ from: 0, instr: 'AUTHOR' }, pAuthor)).toEqual({
			text: 'Katherine Johnson',
		});
		expect(fields.result({ from: 0, instr: '=1/0' }, pDiv)).toBeUndefined();
		expect(
			computedFields(model, { errors: true }).result({ from: 0, instr: '=1/0' }, pDiv),
		).toEqual({ text: ZERO_DIVIDE_ERROR });

		const refreshed = refreshComputedFields(model, { body: true });
		expect((refreshed.blocks[0] as Paragraph).runs[3]?.text).toBe('Katherine Johnson');
		expect((refreshed.blocks[1] as Paragraph).runs[3]?.text).toBe('pass');
		expect((refreshed.blocks[2] as Paragraph).runs[3]?.text).toBe('42');
		expect((refreshed.blocks[3] as Paragraph).runs[3]?.text).toBe('old');
	});
});
