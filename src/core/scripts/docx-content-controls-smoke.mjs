// Fills a template's content controls with the packed package, through every pairing of the
// entries a consumer may combine (ESM and CJS, `ooxml-core/docx` and `ooxml-core/docx/load` to
// load, `ooxml-core/docx` and `ooxml-core/docx/content-controls` to fill), and saves the result
// with the loaded template's `save()`. Each entry is a bundle of its own, so this fails when state
// one writes (the source of each repeated row) is kept where another cannot read it.
import assert from 'node:assert/strict';
import { createRequire } from 'node:module';
import JSZip from 'jszip';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const run = (text) => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;
let id = 10;
const sdt = (tag, content) =>
	`<w:sdt><w:sdtPr><w:id w:val="${id++}"/><w:tag w:val="${tag}"/></w:sdtPr><w:sdtContent>${content}</w:sdtContent></w:sdt>`;
const cell = (inner) =>
	`<w:tc><w:tcPr><w:tcBorders><w:top w:val="double"/></w:tcBorders></w:tcPr><w:p>${inner}</w:p></w:tc>`;
const body =
	sdt('notes', `<w:p>${run('notes')}</w:p>`) +
	`<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>` +
	`<w:tr>${cell(sdt('item_template_name', run('n')))}${cell(sdt('item_template_price', run('p')))}</w:tr>` +
	`</w:tbl><w:p>${run('end')}</w:p>`;
const zip = new JSZip();
zip.file('word/document.xml', `<w:document xmlns:w="${W}"><w:body>${body}</w:body></w:document>`);
const template = await zip.generateAsync({ type: 'uint8array' });
const data = {
	values: { notes: '<p>Hello <b>Ada</b></p><p>Second</p>' },
	html: ['notes'],
	repeat: {
		item: [
			{ name: 'Pen', price: '10' },
			{ name: 'Ink', price: '20' },
		],
	},
};

const require = createRequire(import.meta.url);
const loaders = [(specifier) => import(specifier), async (specifier) => require(specifier)];
let checked = 0;
for (const load of loaders) {
	const docx = await load('ooxml-core/docx');
	const docxLoad = await load('ooxml-core/docx/load');
	const controls = await load('ooxml-core/docx/content-controls');
	for (const open of [docx.loadDocx, docxLoad.loadDocument]) {
		for (const entry of [docx, controls]) {
			const loaded = await open(template);
			const filled = await entry.fillContentControls(loaded.model, data);
			assert.deepEqual(filled.results.map((result) => result.status).sort(), [
				'filled',
				'filled',
				'filled',
				'filled',
				'filled',
				'repeated',
			]);
			const bytes = await loaded.save(filled.model);
			const xml = await (await JSZip.loadAsync(bytes)).file('word/document.xml').async('string');
			assert.equal(xml.match(/<w:tr>/g)?.length, 2, 'every item row is written');
			assert.equal(xml.match(/<w:top w:val="double"\/>/g)?.length, 4, 'copies keep their borders');
			assert.ok(xml.includes('<w:b/>'), 'the HTML is read');
			const reopened = await docx.loadDocx(bytes);
			const values = await entry.readContentControls(reopened.model);
			assert.deepEqual(values, {
				notes: 'Hello Ada\nSecond',
				item_1_name: 'Pen',
				item_1_price: '10',
				item_2_name: 'Ink',
				item_2_price: '20',
			});
			const flat = await entry.flattenContentControls(reopened.model);
			const plain = await (
				await JSZip.loadAsync(await reopened.save(flat))
			)
				.file('word/document.xml')
				.async('string');
			assert.ok(!plain.includes('<w:sdt>'), 'flattened');
			checked++;
		}
	}
}
assert.equal(checked, 8);
