// @vitest-environment jsdom
// Paragraphs with position markers or form fields edited through the editor and saved: the editor
// merges equally formatted runs and keeps or drops marks the way ProseMirror does, so the writer
// sees what a user edit really produces.
import { afterEach, describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { EditorState } from 'prosemirror-state';
import { EditorView } from 'prosemirror-view';
import { loadDocx } from 'ooxml-core/docx';
import { formTextResults } from 'ooxml-core/docx/ui';
import { DocxEditorElement, registerDocxEditor } from './index';
import { docToModel, modelToDoc } from './model-adapter';
import { SearchController } from './search-controller';

registerDocxEditor();

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const run = (text: string) => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;

async function docx(paragraphXml: string): Promise<Uint8Array> {
	const zip = new JSZip();
	zip.file(
		'word/document.xml',
		`<w:document xmlns:w="${W}"><w:body>${paragraphXml}<w:sectPr/></w:body></w:document>`,
	);
	return zip.generateAsync({ type: 'uint8array' });
}

async function open(paragraphXml: string) {
	const loaded = await loadDocx(await docx(paragraphXml));
	const host = document.createElement('div');
	document.body.append(host);
	const view = new EditorView(host, {
		state: EditorState.create({ doc: modelToDoc(loaded.model) }),
	});
	return { loaded, view };
}

async function savedParagraph(
	loaded: Awaited<ReturnType<typeof loadDocx>>,
	view: EditorView,
): Promise<string> {
	const bytes = await loaded.save(docToModel(view.state.doc, loaded.model));
	const xml = await (await JSZip.loadAsync(bytes)).file('word/document.xml')!.async('string');
	return xml.match(/<w:p\b[\s\S]*?<\/w:p>/)?.[0] ?? '';
}

/** Asserts that every needle occurs in `xml`, in this order. */
function expectInOrder(xml: string, needles: string[]) {
	let from = 0;
	for (const needle of needles) {
		const index = xml.indexOf(needle, from);
		expect(index, `${needle} after offset ${from} in ${xml}`).toBeGreaterThanOrEqual(0);
		from = index + needle.length;
	}
}

afterEach(() => document.body.replaceChildren());

describe('saving paragraphs with markers edited in the editor', () => {
	it('keeps a bookmark between the runs of a tracked change that the editor holds as one text', async () => {
		const revision = 'w:id="5" w:author="Ann" w:date="2026-01-01T00:00:00Z"';
		const { loaded, view } = await open(
			`<w:p>${run('Fixed ')}<w:ins ${revision}>${run('left')}<w:bookmarkStart w:id="3" w:name="Mid"/>` +
				`${run('right')}<w:bookmarkEnd w:id="3"/></w:ins>${run(' tail')}</w:p>`,
		);
		const paragraph = view.state.doc.firstChild!;
		expect(paragraph.childCount).toBe(3);
		view.dispatch(view.state.tr.insertText('!', paragraph.nodeSize - 1));
		const xml = await savedParagraph(loaded, view);
		expectInOrder(xml, [
			'<w:ins ',
			'>left<',
			'<w:bookmarkStart w:id="3" w:name="Mid"/>',
			'>right<',
			'<w:bookmarkEnd w:id="3"/>',
			'</w:ins>',
			'> tail!<',
		]);
	});

	it('keeps the markup of a hyperlink whose whole label Replace All replaces', async () => {
		const { loaded, view } = await open(
			`<w:p>${run('See ')}<w:hyperlink w:anchor="t" w:tgtFrame="_blank" w:docLocation="here">` +
				`${run('Docs')}</w:hyperlink>${run(' now')}</w:p>`,
		);
		const search = new SearchController(() => view);
		search.search('Docs');
		expect(search.replaceAll('Help')).toBe(1);
		const xml = await savedParagraph(loaded, view);
		expectInOrder(xml, [
			'>See <',
			'<w:hyperlink w:anchor="t" w:tgtFrame="_blank" w:docLocation="here">',
			'>Help<',
			'</w:hyperlink>',
			'> now<',
		]);
		expect(xml).not.toContain('w:history');
	});
});

describe('saving a form field filled in the editor', () => {
	it('saves text typed into a blank text field as Word writes it, its result still unproofed', async () => {
		const field = (inner: string, props = '') =>
			`<w:r>${props ? `<w:rPr>${props}</w:rPr>` : ''}${inner}</w:r>`;
		// Word's result for a text field with no text: five en spaces, one run each.
		const blank = Array.from({ length: 5 }, () => field('<w:t>\u2002</w:t>', '<w:noProof/>')).join(
			'',
		);
		const editor = document.createElement('docx-editor') as DocxEditorElement;
		document.body.append(editor);
		await editor.load(
			await docx(
				`<w:p>${run('Name: ')}` +
					field(
						'<w:fldChar w:fldCharType="begin"><w:ffData><w:name w:val="Name"/><w:enabled/><w:calcOnExit w:val="0"/><w:textInput/></w:ffData></w:fldChar>',
					) +
					field('<w:instrText xml:space="preserve"> FORMTEXT </w:instrText>') +
					field('<w:fldChar w:fldCharType="separate"/>') +
					blank +
					field('<w:fldChar w:fldCharType="end"/>') +
					'</w:p>',
			),
		);
		const view = (editor as unknown as { view: EditorView }).view;
		const [result] = formTextResults(view.state.doc);
		view.dispatch(view.state.tr.insertText('Ada', result!.from + 2));
		const reopened = (await loadDocx(await editor.saveBytes())).model.blocks[0]!;
		if (reopened.type !== 'paragraph') throw new Error('Expected a paragraph');
		const filled = reopened.runs.filter((item) => item.field?.instr.trim() === 'FORMTEXT');
		expect(filled.map((item) => item.text)).toEqual(['Ada']);
		expect(filled[0]!.sourceRunPropertiesXml).toContain('noProof');
		expect(reopened.runs.find((item) => item.formField)?.formField?.name).toBe('Name');
	});
});
