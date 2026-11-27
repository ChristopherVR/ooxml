// Fills a template's MERGEFIELD fields with the packed package, through every pairing of the
// entries a consumer may combine (ESM and CJS, `ooxml-core/docx` and `ooxml-core/docx/load` to
// load, `ooxml-core/docx` and `ooxml-core/docx/mail-merge` to merge), and saves the result with
// the loaded template's `save()`. Each entry is a bundle of its own, so this fails when state one
// writes (the repeated rows of a region) is kept where another cannot read it.
import assert from 'node:assert/strict';
import { createRequire } from 'node:module';
import JSZip from 'jszip';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const run = (text) => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;
const field = (instruction, cached = '') =>
	`<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> ${instruction} </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>${run(cached)}<w:r><w:fldChar w:fldCharType="end"/></w:r>`;
const cell = (inner) =>
	`<w:tc><w:tcPr><w:tcBorders><w:top w:val="double"/></w:tcBorders></w:tcPr><w:p>${inner}</w:p></w:tc>`;
const body =
	`<w:p>${run('To ')}${field('MERGEFIELD Name', 'x')}</w:p>` +
	`<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid>` +
	`<w:tr>${cell(field('MERGEFIELD TableStart:Items') + field('MERGEFIELD Item', 'x'))}</w:tr>` +
	`<w:tr>${cell(field('MERGEFIELD Price', 'x') + field('MERGEFIELD TableEnd:Items'))}</w:tr>` +
	`</w:tbl><w:p>${run('end')}</w:p>`;
const zip = new JSZip();
zip.file('word/document.xml', `<w:document xmlns:w="${W}"><w:body>${body}</w:body></w:document>`);
const template = await zip.generateAsync({ type: 'uint8array' });
const data = {
	values: { Name: 'Ada' },
	regions: {
		Items: [
			{ Item: 'Pen', Price: '10' },
			{ Item: 'Ink', Price: '20' },
		],
	},
};

const require = createRequire(import.meta.url);
const loaders = [(specifier) => import(specifier), async (specifier) => require(specifier)];
let checked = 0;
for (const load of loaders) {
	const docx = await load('ooxml-core/docx');
	const docxLoad = await load('ooxml-core/docx/load');
	const mailMerge = await load('ooxml-core/docx/mail-merge');
	for (const open of [docx.loadDocx, docxLoad.loadDocument]) {
		for (const entry of [docx, mailMerge]) {
			const loaded = await open(template);
			const { fields, regions } = await entry.listMergeFields(loaded.model);
			assert.deepEqual(fields.map((item) => item.name).sort(), ['Item', 'Name', 'Price']);
			assert.deepEqual(regions, [{ name: 'Items', kind: 'rows', story: 'body' }]);
			const merged = await entry.mergeFields(loaded.model, data, { unlink: true });
			assert.ok(merged.results.every((result) => result.status === 'filled'));
			const bytes = await loaded.save(merged.model);
			const xml = await (await JSZip.loadAsync(bytes)).file('word/document.xml').async('string');
			assert.equal(xml.match(/<w:tr>/g)?.length, 4, 'every region row is written');
			assert.equal(xml.match(/<w:top w:val="double"\/>/g)?.length, 4, 'copies keep their borders');
			assert.ok(!xml.includes('MERGEFIELD'), 'fields are unlinked');
			const reopened = await docx.loadDocx(bytes);
			const texts = [];
			for (const block of reopened.model.blocks) {
				const paragraphs =
					block.type === 'table' ? block.rows.flat().flatMap((item) => item.paragraphs) : [block];
				for (const paragraph of paragraphs) texts.push(paragraph.runs.map((r) => r.text).join(''));
			}
			assert.deepEqual(texts, ['To Ada', 'Pen', '10', 'Ink', '20', 'end']);
			checked++;
		}
	}
}
assert.equal(checked, 8);
