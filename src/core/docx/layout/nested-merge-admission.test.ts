import { readFileSync } from 'node:fs';
import JSZip from 'jszip';
import { expect, it } from 'vitest';
import { loadDocx } from '../index';
import { buildXml, parseXml } from '../../xml';
import { adaptDocumentModel } from './adapter';
import { nestedFragmentEdges, reflowNestedFragments } from './nested-fragment-edges';

const w = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const source = new URL(
	'../__fixtures__/nested-structural-rows/nested-structural-4-short.docx',
	import.meta.url,
);

for (const scope of ['inner', 'enclosing'] as const)
	for (const merge of ['horizontal', 'restart', 'continue'] as const)
		it(`refuses ${scope} ${merge} merge provenance and preserves real source XML`, async () => {
			const zip = await JSZip.loadAsync(readFileSync(source));
			const doc = parseXml(await zip.file('word/document.xml')!.async('string'));
			const target = doc.getElementsByTagNameNS(w, 'tbl')[scope === 'inner' ? 1 : 0]!;
			const grid = target.getElementsByTagNameNS(w, 'tblGrid')[0]!;
			if (merge === 'horizontal') {
				for (const column of Array.from(grid.getElementsByTagNameNS(w, 'gridCol'))) {
					const width = Number(column.getAttributeNS(w, 'w'));
					grid.removeChild(column);
					for (let i = 0; i < 2; i++) {
						const next = doc.createElementNS(w, 'w:gridCol');
						next.setAttributeNS(w, 'w:w', String(width / 2));
						grid.appendChild(next);
					}
				}
				expect(grid.getElementsByTagNameNS(w, 'gridCol')).toHaveLength(4);
			}
			const rows = Array.from(target.childNodes).filter(
				(node): node is Element => node.nodeType === 1 && (node as Element).localName === 'tr',
			);
			for (const row of rows) {
				const cells = Array.from(row.childNodes).filter(
					(node): node is Element => node.nodeType === 1 && (node as Element).localName === 'tc',
				);
				expect(cells).toHaveLength(2);
				for (const cell of cells) {
					const properties = cell.getElementsByTagNameNS(w, 'tcPr')[0]!;
					const marker = doc.createElementNS(w, merge === 'horizontal' ? 'w:gridSpan' : 'w:vMerge');
					marker.setAttributeNS(w, 'w:val', merge === 'horizontal' ? '2' : merge);
					properties.appendChild(marker);
				}
			}
			zip.file('word/document.xml', buildXml(doc));
			const loaded = await loadDocx(await zip.generateAsync({ type: 'uint8array' }));
			const outer = adaptDocumentModel(loaded.model).sections[0]!.blocks.find(
				(block) => block.kind === 'table',
			)!;
			if (outer.kind !== 'table') throw new Error('Missing enclosing table');
			const inner = outer.rows[0]!.cells[0]!.blocks!.find((block) => block.kind === 'table')!;
			if (inner.kind !== 'table') throw new Error('Missing inner table');
			const adapted = scope === 'inner' ? inner : outer;
			expect(adapted.rows.every((row) => row.cells.length === 2)).toBe(true);
			expect(
				adapted.rows[0]!.cells[0]![merge === 'horizontal' ? 'gridSpan' : 'verticalMerge'],
			).toBe(merge === 'horizontal' ? 2 : merge);
			if (scope === 'inner') expect(nestedFragmentEdges(inner)).toBeUndefined();
			expect(reflowNestedFragments(outer.rows[0]!)).toBe(false);
			const bytes = await loaded.save(loaded.model);
			const saved = await JSZip.loadAsync(bytes);
			for (const [name, part] of Object.entries(zip.files)) {
				if (!part.dir)
					expect(await saved.file(name)!.async('uint8array'), name).toEqual(
						await part.async('uint8array'),
					);
			}
			expect((await loadDocx(bytes)).model).toEqual(loaded.model);
		});
