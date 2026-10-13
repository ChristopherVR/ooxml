import { readFileSync } from 'node:fs';
import JSZip from 'jszip';
import { expect, it } from 'vitest';
import { loadDocx } from '../index';
import { buildXml, parseXml } from '../../xml';
import { adaptDocumentModel } from './adapter';
import { nestedFragmentEdges, reflowNestedFragments } from './nested-fragment-edges';
import { tableParagraphKeeps } from './table-paragraph-keeps';

const w = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const source = new URL(
	'../__fixtures__/nested-structural-rows/nested-structural-4-short.docx',
	import.meta.url,
);
for (const scope of ['inner', 'enclosing'] as const)
	for (const rule of ['exact', 'atLeast'] as const)
		it(`refuses explicit zero ${rule} ${scope} row heights without altering source`, async () => {
			const zip = await JSZip.loadAsync(readFileSync(source));
			const doc = parseXml(await zip.file('word/document.xml')!.async('string'));
			const table = doc.getElementsByTagNameNS(w, 'tbl')[scope === 'inner' ? 1 : 0]!;
			const row = table.getElementsByTagNameNS(w, 'tr')[0]!;
			let properties = row.getElementsByTagNameNS(w, 'trPr')[0];
			if (!properties) {
				properties = doc.createElementNS(w, 'w:trPr');
				row.insertBefore(properties, row.firstChild);
			}
			const height = doc.createElementNS(w, 'w:trHeight');
			height.setAttributeNS(w, 'w:val', '0');
			height.setAttributeNS(w, 'w:hRule', rule);
			properties.appendChild(height);
			zip.file('word/document.xml', buildXml(doc));
			const loaded = await loadDocx(await zip.generateAsync({ type: 'uint8array' }));
			const outer = adaptDocumentModel(loaded.model).sections[0]!.blocks.find(
				(block) => block.kind === 'table',
			)!;
			if (outer.kind !== 'table') throw new Error('Missing enclosing table');
			const inner = outer.rows[0]!.cells[0]!.blocks!.find((block) => block.kind === 'table')!;
			if (inner.kind !== 'table') throw new Error('Missing inner table');
			const adapted = scope === 'inner' ? inner : outer;
			expect(adapted.rows[0]!.heightSpecified).toBe(true);
			// The existing numeric layout remains unchanged; source provenance prevents admission.
			expect(adapted.rows[0]!.heightPx).toBeUndefined();
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
			const modelTable = loaded.model.blocks.find((block) => block.type === 'table')!;
			if (modelTable.type !== 'table') throw new Error('Missing enclosing model table');
			// Change a direct enclosing paragraph, leaving both inner and enclosing row properties intact.
			const run = modelTable.rows[0]![0]!.paragraphs.flatMap((paragraph) => paragraph.runs)[0]!;
			run.text = 'Edited ' + run.text;
			const editedBytes = await loaded.save(loaded.model);
			const edited = await JSZip.loadAsync(editedBytes);
			const originalXml = await zip.file('word/document.xml')!.async('string');
			const editedXml = await edited.file('word/document.xml')!.async('string');
			expect(editedXml.match(/<w:trPr\b[^>]*>[\s\S]*?<\/w:trPr>/g)).toEqual(
				originalXml.match(/<w:trPr\b[^>]*>[\s\S]*?<\/w:trPr>/g),
			);
			const editedDoc = parseXml(editedXml);
			expect(buildXml(editedDoc.getElementsByTagNameNS(w, 'tbl')[1]!)).toBe(
				buildXml(doc.getElementsByTagNameNS(w, 'tbl')[1]!),
			);
			for (const [name, part] of Object.entries(zip.files)) {
				if (!part.dir && name !== 'word/document.xml')
					expect(await edited.file(name)!.async('uint8array'), name).toEqual(
						await part.async('uint8array'),
					);
			}
			expect((await loadDocx(editedBytes)).model).toEqual(loaded.model);
		});

it.each(['exact', 'atLeast'])(
	'retains real ordinary-cell zero %s height and refuses keeping',
	async (rule) => {
		const bytes = readFileSync(
			new URL('../__fixtures__/table-keep-lines/table-keep-fresh-on.docx', import.meta.url),
		);
		const zip = await JSZip.loadAsync(bytes);
		const doc = parseXml(await zip.file('word/document.xml')!.async('string'));
		const row = doc.getElementsByTagNameNS(w, 'tr')[0]!;
		const properties = doc.createElementNS(w, 'w:trPr');
		const height = doc.createElementNS(w, 'w:trHeight');
		height.setAttributeNS(w, 'w:val', '0');
		height.setAttributeNS(w, 'w:hRule', rule);
		properties.appendChild(height);
		row.insertBefore(properties, row.firstChild);
		zip.file('word/document.xml', buildXml(doc));
		const loaded = await loadDocx(await zip.generateAsync({ type: 'uint8array' }));
		const model = loaded.model.blocks.find((block) => block.type === 'table')!;
		if (model.type !== 'table') throw new Error('Missing ordinary model table');
		expect(model.rowProperties![0]).toEqual({ heightTwips: 0, heightRule: rule });
		const table = adaptDocumentModel(loaded.model).sections[0]!.blocks.find(
			(block) => block.kind === 'table',
		)!;
		if (table.kind !== 'table') throw new Error('Missing ordinary layout table');
		expect(table.rows[0]!.heightSpecified).toBe(true);
		expect(table.rows[0]!.heightPx).toBeUndefined();
		expect(tableParagraphKeeps(table)).toBeUndefined();
		const savedBytes = await loaded.save(loaded.model);
		const saved = await JSZip.loadAsync(savedBytes);
		for (const [name, part] of Object.entries(zip.files)) {
			if (!part.dir)
				expect(await saved.file(name)!.async('uint8array'), name).toEqual(
					await part.async('uint8array'),
				);
		}
		expect((await loadDocx(savedBytes)).model).toEqual(loaded.model);
	},
);
