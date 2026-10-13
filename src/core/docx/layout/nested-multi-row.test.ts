import { readFileSync } from 'node:fs';
import { expect, it } from 'vitest';
import JSZip from 'jszip';
import { loadDocx } from '../index';
import { layoutDocumentModel } from './layout';
import { parseXml, buildXml } from '../../xml';
const twoRoot = new URL('../__fixtures__/nested-multi-row/', import.meta.url);
const evidence = JSON.parse(readFileSync(new URL('native-measurements.json', twoRoot), 'utf8')) as {
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
const threeRoot = new URL('../__fixtures__/nested-three-row/', import.meta.url);
const threeEvidence = JSON.parse(
	readFileSync(new URL('native-measurements.json', threeRoot), 'utf8'),
) as typeof evidence;
for (const { native, root } of [
	...evidence.cases.map((native) => ({ native, root: twoRoot })),
	...threeEvidence.cases.map((native) => ({ native, root: threeRoot })),
])
	it('matches native multi-row physical bands and preserves source: ' + native.name, async () => {
		const bytes = new Uint8Array(readFileSync(new URL(native.name + '.docx', root))),
			loaded = await loadDocx(bytes);
		const layout = layoutDocumentModel(loaded.model, measure);
		expect(layout.pages).toHaveLength(native.pages);
		for (const [index, page] of layout.pages.entries()) {
			const table = page.columns[0]!.blocks.find((b) => b.kind === 'table');
			if (!table || table.kind !== 'table') throw new Error('Expected enclosing table');
			const row = table.rows[0]!,
				frame = row.cellTables![0]![0]!,
				sample = native.measurements[index]!;
			const lines = row.cells.flatMap((cell) =>
				cell.flatMap((p) =>
					p.lines.map((line) => ({
						text: line.fragments.map((f) => f.text).join(''),
						top: (p.yPx + line.yPx) * 0.75,
					})),
				),
			);
			expect(
				lines
					.filter((l) => /^T[012]/.test(l.text))
					.map((l) => l.text)
					.sort(),
			).toEqual(
				sample.lines
					.filter((l) => /^T[012]/.test(l.text))
					.map((l) => l.text.trim())
					.sort(),
			);
			const top = Math.min(...sample.borders.map((b) => b.rectPt[1]!)),
				bottom = Math.max(...sample.borders.map((b) => b.rectPt[3]!));
			const left = Math.min(...sample.borders.map((b) => b.rectPt[0]!)),
				right = Math.max(...sample.borders.map((b) => b.rectPt[2]!));
			expect(Math.abs(frame.heightPx * 0.75 - (bottom - top))).toBeLessThan(0.2);
			expect(Math.abs(frame.widthPx! * 0.75 - (right - left))).toBeLessThan(0.1);
			expect(frame.rows[0]!.geometry![0]!.borders!.top!.widthPx * 0.75).toBe(2);
			expect(frame.rows.at(-1)!.geometry![0]!.borders!.bottom!.widthPx * 0.75).toBe(0.5);
			const first = lines.find((l) => l.text.startsWith('T1'))!,
				nativeFirst = sample.lines.find((l) => l.text.trim() === first.text)!;
			expect(
				Math.abs(first.top - frame.yPx * 0.75 - (nativeFirst.baselinePt - 9.6 - top)),
			).toBeLessThan(0.15);
			const outer = lines.find((l) => l.text === 'T2R0C0L00');
			if (outer)
				expect(
					Math.abs(
						outer.top -
							first.top -
							(sample.lines.find((l) => l.text.startsWith('T2'))!.baselinePt -
								nativeFirst.baselinePt),
					),
				).toBeLessThan(0.2);
		}
		const original = await JSZip.loadAsync(bytes),
			noOp = await JSZip.loadAsync(await loaded.save(loaded.model));
		for (const [path, file] of Object.entries(original.files))
			if (!file.dir)
				expect(await noOp.file(path)!.async('uint8array')).toEqual(await file.async('uint8array'));
		expect((await loadDocx(await loaded.save(loaded.model))).model).toEqual(loaded.model);
		const outer = loaded.model.blocks.find((b) => b.type === 'table');
		if (!outer || outer.type !== 'table') throw new Error('Missing source table');
		const cell = outer.rows[0]![0]!,
			run = cell.paragraphs[0]!.runs.find((r) => r.text === 'T2R0C0L00')!;
		run.text = 'E2R0C0L00';
		const edited = await loaded.save(loaded.model),
			saved = await JSZip.loadAsync(edited);
		const nested = (xml: string) => {
			const doc = parseXml(xml);
			return buildXml(
				Array.from(
					doc.getElementsByTagNameNS(
						'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
						'tbl',
					),
				)[1]!,
			);
		};
		expect(nested(await saved.file('word/document.xml')!.async('string'))).toBe(
			nested(await original.file('word/document.xml')!.async('string')),
		);
		for (const [path, file] of Object.entries(original.files))
			if (!file.dir && path !== 'word/document.xml')
				expect(await saved.file(path)!.async('uint8array')).toEqual(await file.async('uint8array'));
		expect(layoutDocumentModel((await loadDocx(edited)).model, measure).pages).toHaveLength(
			native.pages,
		);
	});
