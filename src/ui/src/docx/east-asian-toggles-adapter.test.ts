// @vitest-environment jsdom
import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { loadDocx, saveDocx, type DocumentModel, type Paragraph } from 'ooxml-core/docx';
import { adaptDocumentModel } from 'ooxml-core/docx/layout';
import { docToModel, modelToDoc } from './model-adapter';
import { schema } from './schema';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const BODY =
	'<w:p><w:pPr><w:kinsoku w:val="0"/><w:autoSpaceDE w:val="0"/><w:autoSpaceDN w:val="0"/><w:overflowPunct w:val="0"/></w:pPr><w:r><w:rPr><w:lang w:eastAsia="ja-JP"/></w:rPr><w:t>一。</w:t></w:r></w:p>' +
	'<w:p><w:r><w:rPr><w:lang w:eastAsia="ja-JP"/></w:rPr><w:t>二。</w:t></w:r></w:p>';

async function packageBytes(): Promise<Uint8Array> {
	const zip = new JSZip();
	zip.file(
		'word/document.xml',
		`<w:document xmlns:w="${W}"><w:body>${BODY}<w:sectPr/></w:body></w:document>`,
	);
	return zip.generateAsync({ type: 'uint8array' });
}

/** Loads the model into the editor, types into each paragraph and converts it back. */
function editInEditor(model: DocumentModel): DocumentModel {
	const doc = modelToDoc(model);
	const paragraphs: ReturnType<typeof schema.nodes.paragraph.create>[] = [];
	doc.forEach((node) => {
		paragraphs.push(node.type.create(node.attrs, schema.text(`${node.textContent}追加`)));
	});
	return docToModel(schema.nodes.doc.create(doc.attrs, paragraphs), model, doc);
}

const layoutSwitches = (model: DocumentModel) =>
	adaptDocumentModel(model).sections[0]!.blocks.map((block) =>
		block.kind === 'paragraph'
			? [block.kinsoku, block.autoSpaceDE, block.autoSpaceDN, block.overflowPunct]
			: [],
	);

describe('East Asian typography toggles through the editor', () => {
	it('keeps w:kinsoku and the spacing toggles off after an edit, a save and a reload', async () => {
		const loaded = await loadDocx(await packageBytes());
		const edited = editInEditor(loaded.model);
		const first = edited.blocks[0] as Paragraph;
		expect(first.runs.map((run) => run.text).join('')).toBe('一。追加');
		expect([first.kinsoku, first.autoSpaceDE, first.autoSpaceDN, first.overflowPunct]).toEqual([
			false,
			false,
			false,
			false,
		]);
		expect(layoutSwitches(edited)[0]).toEqual([false, false, false, false]);

		const saved = await saveDocx(edited);
		const xml = await (await JSZip.loadAsync(saved)).file('word/document.xml')!.async('string');
		expect(xml).toContain('<w:kinsoku w:val="0"/>');
		// The paragraph without the elements saves without them.
		expect(xml.match(/<w:kinsoku/g)).toHaveLength(1);
		expect(xml.match(/<w:overflowPunct/g)).toHaveLength(1);

		const reloaded = await loadDocx(saved);
		expect((reloaded.model.blocks[0] as Paragraph).runs.map((run) => run.text).join('')).toBe(
			'一。追加',
		);
		expect(layoutSwitches(reloaded.model)).toEqual([
			[false, false, false, false],
			[undefined, undefined, undefined, undefined],
		]);
	});
});
