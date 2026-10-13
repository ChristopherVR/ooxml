import { readFileSync } from 'node:fs';
import JSZip from 'jszip';
import { expect, it } from 'vitest';
import { loadDocx } from '../index';
import { buildXml, parseXml } from '../../xml';
import { adaptDocumentModel } from './adapter';
import { layoutDocumentModel } from './layout';
import { layoutBorderedRow } from './table-border-layout';
import { nestedFrameProfile } from './nested-frame-profile';
import { tableParagraphKeeps } from './table-paragraph-keeps';

const w = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const evidence = JSON.parse(
	readFileSync(
		new URL('../__fixtures__/table-keep-boundaries/native-boundaries.json', import.meta.url),
		'utf8',
	),
) as {
	cases: {
		name: string;
		borderPt: number;
		paddingPt: number;
		mirror: boolean;
		keep: boolean;
		remainingPt: number;
		minimumPt: number;
		pages: {
			body: string[];
			topPt: number | null;
			bottomPt: number | null;
			firstBaselinePt: number | null;
		}[];
	}[];
};
const measure = { widthOf: (text: string) => text.length * 5, lineHeightOf: () => 16 };
for (const native of evidence.cases)
	it('retains native whole-row startup at the fit boundary: ' + native.name, async () => {
		const filename =
			'table-keep-' +
			(native.mirror ? 'mirror-fresh' : 'fresh') +
			'-' +
			(native.keep ? 'on' : 'off') +
			'.docx';
		const zip = await JSZip.loadAsync(
			readFileSync(new URL('../__fixtures__/table-keep-lines/' + filename, import.meta.url)),
		);
		const doc = parseXml(await zip.file('word/document.xml')!.async('string'));
		const body = doc.getElementsByTagNameNS(w, 'body')[0]!;
		const paragraph = Array.from(body.childNodes).find(
			(node) => node.nodeType === 1 && (node as Element).localName === 'p',
		) as Element;
		const runs = Array.from(paragraph.getElementsByTagNameNS(w, 'r'));
		for (const run of runs.slice(1)) paragraph.removeChild(run);
		paragraph
			.getElementsByTagNameNS(w, 'spacing')[0]!
			.setAttributeNS(w, 'w:line', String(Math.round((220 - native.remainingPt) * 20)));
		const sourceTable = doc.getElementsByTagNameNS(w, 'tbl')[0]!;
		for (const tag of ['tcBorders', 'tblBorders'])
			for (const borders of Array.from(sourceTable.getElementsByTagNameNS(w, tag))) {
				for (const edge of Array.from(borders.childNodes))
					if (edge.nodeType === 1)
						(edge as Element).setAttributeNS(w, 'w:sz', String(Math.round(native.borderPt * 8)));
			}
		for (const cell of Array.from(sourceTable.getElementsByTagNameNS(w, 'tc'))) {
			const properties = cell.getElementsByTagNameNS(w, 'tcPr')[0]!;
			let margins = properties.getElementsByTagNameNS(w, 'tcMar')[0];
			if (!margins) {
				margins = doc.createElementNS(w, 'w:tcMar');
				properties.appendChild(margins);
			}
			for (const side of ['top', 'bottom']) {
				let item = margins.getElementsByTagNameNS(w, side)[0];
				if (!item) {
					item = doc.createElementNS(w, 'w:' + side);
					margins.appendChild(item);
				}
				item.setAttributeNS(w, 'w:w', String(Math.round(native.paddingPt * 20)));
				item.setAttributeNS(w, 'w:type', 'dxa');
			}
		}
		zip.file('word/document.xml', buildXml(doc));
		const loaded = await loadDocx(await zip.generateAsync({ type: 'uint8array' }));
		const table = adaptDocumentModel(loaded.model).sections[0]!.blocks.find(
			(block) => block.kind === 'table',
		)!;
		if (table.kind !== 'table') throw new Error('Missing table');
		const row = layoutBorderedRow(
			table.rows[0]!,
			table.widthPx!,
			measure,
			() => {},
			nestedFrameProfile(table),
		);
		if (native.keep)
			expect(tableParagraphKeeps(table)!.startupMinimum(row) * 0.75).toBe(native.minimumPt);
		const result = layoutDocumentModel(loaded.model, measure);
		expect(result.pages).toHaveLength(native.pages.length);
		for (const [index, page] of result.pages.entries()) {
			const sample = native.pages[index]!;
			const physical = page.columns[0]!.blocks.find((block) => block.kind === 'table');
			const actual =
				physical?.kind === 'table'
					? physical.rows.flatMap((row) =>
							row.cells.flatMap((cell) =>
								cell.flatMap((p) =>
									p.lines.map((line) => ({
										text: line.fragments.map((f) => f.text).join(''),
										top: (page.marginTopPx + physical.yPx + row.yPx + p.yPx + line.yPx) * 0.75,
									})),
								),
							),
						)
					: [];
			expect(actual.map((line) => line.text).sort()).toEqual([...sample.body].sort());
			if (sample.topPt === null) expect(physical).toBeUndefined();
			else {
				if (!physical || physical.kind !== 'table') throw new Error('Missing physical frame');
				expect(Math.abs((page.marginTopPx + physical.yPx) * 0.75 - sample.topPt)).toBeLessThan(0.2);
				expect(Math.abs(physical.heightPx * 0.75 - (sample.bottomPt! - sample.topPt))).toBeLessThan(
					0.2,
				);
				expect(Math.abs(actual[0]!.top - (sample.firstBaselinePt! - 9.6))).toBeLessThan(0.15);
			}
		}
		const savedBytes = await loaded.save(loaded.model);
		const saved = await JSZip.loadAsync(savedBytes);
		for (const [name, part] of Object.entries(zip.files))
			if (!part.dir)
				expect(await saved.file(name)!.async('uint8array'), name).toEqual(
					await part.async('uint8array'),
				);
		expect((await loadDocx(savedBytes)).model).toEqual(loaded.model);
	});
