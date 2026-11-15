import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { loadDocx } from './index';
import {
	acceptAllRevisions,
	acceptRevision,
	listRevisions,
	rejectAllRevisions,
	rejectRevision,
} from './revision-commands';
import { numberRevisionIds } from './revision-ids';
import { expectTable } from './test-support/access';

const ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const row = (text: string, trPr = '') =>
	`<w:tr>${trPr}<w:tc><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:tc></w:tr>`;
const document = (rows: string) =>
	`<w:document xmlns:w="${ns}"><w:body><w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid>${rows}</w:tbl><w:p/><w:sectPr/></w:body></w:document>`;
const TRACKED = document(
	row('Kept') +
		row(
			'Added',
			'<w:trPr><w:ins w:id="7" w:author="Ada" w:date="2026-01-01T00:00:00Z"/></w:trPr>',
		) +
		row('Removed', '<w:trPr><w:del w:id="8" w:author="Ada"/></w:trPr>'),
);

async function load(xml: string) {
	const zip = new JSZip();
	zip.file('word/document.xml', xml);
	return loadDocx(await zip.generateAsync({ type: 'uint8array' }));
}
const rowTexts = (model: Awaited<ReturnType<typeof load>>['model']) =>
	expectTable(model.blocks[0]).rows.map((cells) => cells[0]!.paragraphs[0]!.runs[0]?.text);

describe('tracked table rows (w:trPr/w:ins, w:trPr/w:del)', () => {
	it('reads row insertions and deletions and lists them before the row text', async () => {
		const { model } = await load(TRACKED);
		const table = expectTable(model.blocks[0]);
		expect(table.rowProperties?.map((props) => props.revision?.kind)).toEqual([
			undefined,
			'insert',
			'delete',
		]);
		expect(table.rowProperties?.[1]?.revision).toMatchObject({
			id: '7',
			author: 'Ada',
			date: '2026-01-01T00:00:00Z',
		});
		expect(listRevisions(model).map((entry) => [entry.kind, entry.row?.index])).toEqual([
			['insert', 1],
			['delete', 2],
		]);
	});

	it('accepts and rejects whole rows', async () => {
		const { model } = await load(TRACKED);
		expect(rowTexts(acceptAllRevisions(model))).toEqual(['Kept', 'Added']);
		expect(rowTexts(rejectAllRevisions(model))).toEqual(['Kept', 'Removed']);
		expect(rowTexts(acceptRevision(model, '8'))).toEqual(['Kept', 'Added']);
		expect(rowTexts(rejectRevision(model, '7'))).toEqual(['Kept', 'Removed']);
		const kept = expectTable(acceptRevision(model, '7').blocks[0]);
		expect(kept.rowProperties?.[1]?.revision).toBeUndefined();
	});

	it('writes a new row revision and clears an accepted one', async () => {
		const loaded = await load(document(row('One') + row('Two')));
		const table = expectTable(loaded.model.blocks[0]);
		table.rowProperties = [{}, { revision: { kind: 'delete', author: 'Ada', id: 'new-row' } }];
		const xml = await (
			await JSZip.loadAsync(await loaded.save())
		)
			.file('word/document.xml')!
			.async('string');
		expect(xml).toMatch(/<w:trPr><w:del w:id="\d+" w:author="Ada"\/><\/w:trPr>/);
		const reloaded = await loadDocx(await loaded.save());
		const accepted = acceptAllRevisions(reloaded.model);
		const saved = await (
			await JSZip.loadAsync(await reloaded.save(accepted))
		)
			.file('word/document.xml')!
			.async('string');
		expect(saved).not.toContain('<w:del ');
		expect(saved.match(/<w:tr[ >]/g)).toHaveLength(1);
	});

	it('numbers editor row revision ids as decimals', async () => {
		const { model } = await load(document(row('One')));
		expectTable(model.blocks[0]).rowProperties = [
			{ revision: { kind: 'insert', author: 'Ada', id: 'editor-id' } },
		];
		const numbered = numberRevisionIds(model, 4);
		expect(expectTable(numbered.blocks[0]).rowProperties?.[0]?.revision?.id).toBe('5');
		expect(expectTable(model.blocks[0]).rowProperties?.[0]?.revision?.id).toBe('editor-id');
	});

	it('keeps a cell end mark when its tracked insertion is rejected', async () => {
		const { model } = await load(
			document(
				'<w:tr><w:tc><w:p><w:pPr><w:rPr><w:ins w:id="3" w:author="Ada"/></w:rPr></w:pPr><w:r><w:t>Cell</w:t></w:r></w:p></w:tc></w:tr>',
			),
		);
		expect(rowTexts(rejectAllRevisions(model))).toEqual(['Cell']);
	});

	it('restructures tables whose rows carry Word row property exceptions (w:tblPrEx)', async () => {
		const exceptions =
			'<w:tblPrEx><w:tblCellMar><w:top w:w="0" w:type="dxa"/></w:tblCellMar></w:tblPrEx>';
		const loaded = await load(
			document(`<w:tr>${exceptions}<w:tc><w:p><w:r><w:t>One</w:t></w:r></w:p></w:tc></w:tr>`),
		);
		const table = expectTable(loaded.model.blocks[0]);
		expect(table.structureEditable).toBe(true);
		table.rows.push([
			{ paragraphs: [{ type: 'paragraph', id: 'added', runs: [{ text: 'Two' }] }] },
		]);
		table.rowProperties = [{ header: true }, {}];
		const xml = await (
			await JSZip.loadAsync(await loaded.save())
		)
			.file('word/document.xml')!
			.async('string');
		expect(xml.match(/<w:tr[ >]/g)).toHaveLength(2);
		expect(xml).toContain(`<w:tr>${exceptions}<w:trPr><w:tblHeader/></w:trPr><w:tc>`);
	});
});
