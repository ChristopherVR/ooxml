import { afterEach, expect, it, vi } from 'vitest';
import { isFieldLocked, loadDocx, type DocumentModel, type Paragraph, type TextRun } from './index';
import { fieldDisplayText } from './layout/page-fields';
import { adaptDocumentModel } from './layout/adapter';
import { fieldLockFixture, fieldLockCases } from './test-support/field-lock-fixture';
import { parseBlocksFromContainer } from './block-parser';
import { parseXml, WORD_NS } from './xml';

afterEach(() => vi.useRealTimers());

it.each([
	[undefined, false],
	[{}, false],
	[{ dirty: true }, false],
	[{ locked: false, dirty: true }, false],
	[{ locked: true, dirty: true }, true],
])('checks only the explicit lock flag %j', (fieldFlags, expected) => {
	expect(isFieldLocked(fieldFlags === undefined ? {} : { fieldFlags })).toBe(expected);
	expect(isFieldLocked(undefined)).toBe(false);
});

const values = {
	page: '4',
	numPages: '8',
	sectionPages: '3',
	now: new Date(2026, 0, 2, 13, 14, 15),
};

it.each(['PAGE', 'NUMPAGES', 'SECTIONPAGES', 'DATE', 'TIME'])(
	'preserves locked %s cached text, including an empty cache, while dirty',
	(instr) => {
		for (const text of ['Saved cache', '']) {
			const run: TextRun = {
				text,
				field: { instr },
				fieldFlags: { locked: true, dirty: true },
			};
			expect(fieldDisplayText(run, values)).toBe(text);
			expect(run.fieldFlags).toEqual({ locked: true, dirty: true });
		}
	},
);

it.each(['PAGE', 'NUMPAGES', 'SECTIONPAGES', 'DATE', 'TIME'])(
	'updates explicitly unlocked dirty %s results',
	(instr) => {
		const run: TextRun = {
			text: 'Saved cache',
			field: { instr },
			fieldFlags: { locked: false, dirty: true },
		};
		expect(fieldDisplayText(run, values)).not.toBe('Saved cache');
		expect(run.fieldFlags).toEqual({ locked: false, dirty: true });
	},
);

it.each(['DATE', 'TIME'])('honors %s locks in body layout without mutating the cache', (instr) => {
	vi.useFakeTimers();
	vi.setSystemTime(values.now);
	const runs: TextRun[] = [
		{ text: 'Saved cache', field: { instr }, fieldFlags: { locked: true, dirty: true } },
		{ text: 'Saved cache', field: { instr }, fieldFlags: { locked: false, dirty: true } },
	];
	const model: DocumentModel = {
		blocks: [{ type: 'paragraph', id: 'p1', runs }],
		page: {
			width: 816,
			height: 1056,
			marginTop: 96,
			marginRight: 96,
			marginBottom: 96,
			marginLeft: 96,
		},
		warnings: [],
	};
	const snapshot = structuredClone(model);
	const paragraph = adaptDocumentModel(model).sections[0]!.blocks[0]!;
	expect(paragraph).toMatchObject({
		runs: [{ text: 'Saved cache' }, { text: fieldDisplayText(runs[1]!, values) }],
	});
	expect(model).toEqual(snapshot);
});

it.each(['simple', 'complex'] as const)(
	'uses imported %s result lock metadata after save/reload',
	async (kind) => {
		const loaded = await loadDocx(await fieldLockFixture(kind));
		const reloaded = await loadDocx(await loaded.save(loaded.model));
		for (const model of [loaded.model, reloaded.model]) {
			const results = (model.blocks as Paragraph[])
				.slice(1)
				.map((block) => block.runs.find((run) => run.field)!);
			expect(results.map(isFieldLocked)).toEqual(fieldLockCases.map((item) => item.locked));
			expect(results.map((run) => run.text)).toEqual(fieldLockCases.map((item) => item.cache));
		}
	},
);

it('isolates inner unlocked results and resumes the outer lock for display', () => {
	const marker = (type: string, flags = '') =>
		`<w:r><w:fldChar w:fldCharType="${type}" ${flags}/></w:r>`;
	const instruction = '<w:r><w:instrText> PAGE </w:instrText></w:r>';
	const cache = (text: string) => `<w:r><w:t>${text}</w:t></w:r>`;
	const xml = parseXml(
		`<w:body xmlns:w="${WORD_NS}"><w:p>${marker('begin', 'w:fldLock="true" w:dirty="true"')}${instruction}${marker('separate')}${cache('outer-before')}${marker('begin', 'w:fldLock="false"')}${instruction}${marker('separate')}${cache('inner')}${marker('end')}${cache('outer-after')}${marker('end')}</w:p></w:body>`,
	);
	const runs = (parseBlocksFromContainer(xml.documentElement)[0] as Paragraph).runs.filter(
		(run) => run.field,
	);
	expect(runs.map((run) => fieldDisplayText(run, values))).toEqual([
		'outer-before',
		'4',
		'outer-after',
	]);
	expect(runs.map(isFieldLocked)).toEqual([true, false, true]);
});

it.each([
	[undefined, { locked: false }],
	[
		{ locked: false, dirty: true },
		{ locked: true, dirty: true },
	],
	[
		{ locked: true, dirty: false },
		{ locked: true, dirty: true },
	],
] as const)(
	'exports anonymous adjacent simple fields with unequal flags independently (%j, %j)',
	async (first, second) => {
		const loaded = await loadDocx(await fieldLockFixture('simple'));
		const paragraph = loaded.model.blocks[0] as Paragraph;
		const runs: TextRun[] = [first, second].map((fieldFlags, index) => ({
			text: String(index),
			field: { instr: 'REF Target', simple: true },
			...(fieldFlags && { fieldFlags }),
		}));
		const bytes = await loaded.save({ ...loaded.model, blocks: [{ ...paragraph, runs }] });
		const reloaded = await loadDocx(bytes);
		const results = (reloaded.model.blocks[0] as Paragraph).runs;
		expect(results.map((run) => run.text)).toEqual(['0', '1']);
		expect(results.map((run) => run.fieldFlags)).toEqual([first, second]);
		expect(new Set(results.map((run) => run.fieldInstanceId)).size).toBe(2);
	},
);
