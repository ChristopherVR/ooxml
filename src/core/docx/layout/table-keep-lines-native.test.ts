import { readFileSync } from 'node:fs';
import { expect, it } from 'vitest';
import JSZip from 'jszip';
import { loadDocx, twips } from '../index';
import { layoutDocumentModel } from './layout';
import { adaptDocumentModel } from './adapter';
import { tableParagraphKeeps } from './table-paragraph-keeps';

const root = new URL('../__fixtures__/table-keep-lines/', import.meta.url);
const evidence = JSON.parse(readFileSync(new URL('native-measurements.json', root), 'utf8')) as {
	cases: {
		name: string;
		pages: number;
		measurements: {
			borders: { rectPt: number[] }[];
			lines: { text: string; leftPt: number; baselinePt: number }[];
		}[];
	}[];
};
const measure = { widthOf: (text: string) => text.length * 5, lineHeightOf: () => 16 };
const cacheNormalized = (xml: string) =>
	xml
		.replace(/\r\n/g, '\n')
		.replace(/<w:p\b[^>]*>[\s\S]*?<\/w:p>/g, (paragraph) =>
			paragraph.includes('T1R0C0LE00')
				? paragraph
						.replace(
							/ xmlns:w="http:\/\/schemas.openxmlformats.org\/wordprocessingml\/2006\/main"/g,
							'',
						)
						.replace(/<w:lastRenderedPageBreak\b[^>]*\/>/g, '')
				: paragraph,
		);

for (const native of evidence.cases) {
	it('matches native ordinary-cell keeping, page bands and ink: ' + native.name, async () => {
		const loaded = await loadDocx(
			new Uint8Array(readFileSync(new URL(native.name + '.docx', root))),
		);
		const result = layoutDocumentModel(loaded.model, measure);
		expect(result.pages).toHaveLength(native.pages);
		for (const [index, page] of result.pages.entries()) {
			const sample = native.measurements[index]!,
				column = page.columns[0]!;
			const actual = column.blocks.flatMap((block) =>
				block.kind === 'paragraph'
					? block.lines.map((line) => ({
							text: line.fragments.map((f) => f.text).join(''),
							top: (page.marginTopPx + block.yPx + line.yPx) * 0.75,
						}))
					: block.rows.flatMap((row) =>
							row.cells.flatMap((cell) =>
								cell.flatMap((p) =>
									p.lines.map((line) => ({
										text: line.fragments.map((f) => f.text).join(''),
										top: (page.marginTopPx + block.yPx + row.yPx + p.yPx + line.yPx) * 0.75,
									})),
								),
							),
						),
			);
			expect(
				actual
					.filter((line) => /^T[01]/.test(line.text))
					.map((line) => line.text)
					.sort(),
			).toEqual(
				sample.lines
					.filter((line) => /^T[01]/.test(line.text))
					.map((line) => line.text.trim())
					.sort(),
			);
			for (const line of actual.filter((line) => /^T[01]/.test(line.text))) {
				const reference = sample.lines.find((candidate) => candidate.text.trim() === line.text)!;
				expect(Math.abs(line.top - (reference.baselinePt - 9.6))).toBeLessThan(0.2);
			}
			const table = column.blocks.find((block) => block.kind === 'table');
			if (!sample.borders.length) expect(table).toBeUndefined();
			else {
				if (!table || table.kind !== 'table') throw new Error('Missing native physical row');
				const top = Math.min(...sample.borders.map((b) => b.rectPt[1]!));
				const bottom = Math.max(...sample.borders.map((b) => b.rectPt[3]!));
				expect(Math.abs(table.heightPx * 0.75 - (bottom - top))).toBeLessThan(0.2);
				const first = actual.find((line) => /^T1/.test(line.text))!;
				const nativeFirst = sample.lines.find((line) => line.text.trim() === first.text)!;
				expect(Math.abs(first.top - (nativeFirst.baselinePt - 9.6))).toBeLessThan(0.15);
			}
		}
	});
	it('preserves no-op source and separates edited cache effects: ' + native.name, async () => {
		const bytes = new Uint8Array(readFileSync(new URL(native.name + '.docx', root)));
		const loaded = await loadDocx(bytes),
			original = await JSZip.loadAsync(bytes);
		const noOpBytes = await loaded.save(loaded.model),
			noOp = await JSZip.loadAsync(noOpBytes);
		for (const [path, file] of Object.entries(original.files))
			if (!file.dir)
				expect(await noOp.file(path)!.async('uint8array')).toEqual(await file.async('uint8array'));
		expect((await loadDocx(noOpBytes)).model).toEqual(loaded.model);
		const table = loaded.model.blocks.find((block) => block.type === 'table');
		if (!table || table.type !== 'table') throw new Error('Missing ordinary table');
		const paragraph = table.rows[0]!.flatMap((cell) => cell.paragraphs).find((p) =>
			p.runs.some((run) => run.text === 'T1R0C0LK00'),
		)!;
		paragraph.runs.find((run) => run.text === 'T1R0C0LK00')!.text = 'T1R0C0LE00';
		const editedBytes = await loaded.save(loaded.model),
			edited = await JSZip.loadAsync(editedBytes);
		const expectedXml = (await original.file('word/document.xml')!.async('string')).replace(
			'T1R0C0LK00',
			'T1R0C0LE00',
		);
		expect(cacheNormalized(await edited.file('word/document.xml')!.async('string'))).toBe(
			cacheNormalized(expectedXml),
		);
		for (const [path, file] of Object.entries(original.files))
			if (!file.dir && path !== 'word/document.xml')
				expect(await edited.file(path)!.async('uint8array')).toEqual(
					await file.async('uint8array'),
				);
		const reopened = await loadDocx(editedBytes);
		const reopenedTable = reopened.model.blocks.find((block) => block.type === 'table');
		if (!reopenedTable || reopenedTable.type !== 'table') throw new Error('Missing reopened table');
		expect(
			reopenedTable.rows[0]!.flatMap((cell) => cell.paragraphs).find((p) =>
				p.runs.some((run) => run.text === 'T1R0C0LE00'),
			)!,
		).toEqual(paragraph);
		const expected = structuredClone(loaded.model);
		if (table.structureEditable === false && reopenedTable.structureEditable === true) {
			// Existing writer clears stale page-break cache; source-admission diagnostics are derived.
			expect(expectedXml).toContain('lastRenderedPageBreak');
			const expectedTable = expected.blocks.find((block) => block.type === 'table');
			if (!expectedTable || expectedTable.type !== 'table')
				throw new Error('Missing expected table');
			expectedTable.structureEditable = true;
			expected.warnings = expected.warnings.filter(
				(warning) =>
					warning !==
					'Merged, nested, or complex tables can be read, but their row and column structure cannot be edited safely.',
			);
		}
		expect(reopened.model).toEqual(expected);
		expect(layoutDocumentModel(reopened.model, measure).pages).toHaveLength(native.pages);
	});
}
it.each(['horizontal', 'vertical'] as const)(
	'retains source merge provenance for keep refusal: %s',
	async (kind) => {
		const loaded = await loadDocx(
			new Uint8Array(readFileSync(new URL('table-keep-partial-on.docx', root))),
		);
		const table = loaded.model.blocks.find((block) => block.type === 'table');
		if (!table || table.type !== 'table') throw new Error('Missing source table');
		if (kind === 'horizontal') table.rows[0]![0]!.gridSpan = 2;
		else table.rows[0]![0]!.verticalMerge = 'restart';
		const input = adaptDocumentModel(loaded.model).sections[0]!.blocks.find(
			(block) => block.kind === 'table',
		);
		if (!input || input.kind !== 'table') throw new Error('Missing adapted table');
		expect(input.rows[0]!.cells[0]).toMatchObject(
			kind === 'horizontal' ? { gridSpan: 2 } : { verticalMerge: 'restart' },
		);
		expect(tableParagraphKeeps(input)).toBeUndefined();
	},
);
it('refuses explicit zero source heights without changing the existing box-height adapter', async () => {
	const loaded = await loadDocx(
		new Uint8Array(readFileSync(new URL('table-keep-partial-on.docx', root))),
	);
	const table = loaded.model.blocks.find((block) => block.type === 'table');
	if (!table || table.type !== 'table') throw new Error('Missing source table');
	table.rowProperties = [{ heightTwips: twips(0), heightRule: 'exact' }];
	const input = adaptDocumentModel(loaded.model).sections[0]!.blocks.find(
		(block) => block.kind === 'table',
	);
	if (!input || input.kind !== 'table') throw new Error('Missing adapted table');
	expect(input.rows[0]!.heightSpecified).toBe(true);
	expect(input.rows[0]!.heightPx).toBeUndefined();
	expect(tableParagraphKeeps(input)).toBeUndefined();
});
