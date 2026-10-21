// @vitest-environment jsdom
// Paragraphs with position markers edited through the editor and saved: the editor merges equally
// formatted runs and keeps or drops marks the way ProseMirror does, so the writer sees what a user
// edit really produces.
import { afterEach, describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { EditorState } from 'prosemirror-state';
import { EditorView } from 'prosemirror-view';
import { loadDocx } from 'ooxml-core/docx';
import { docToModel, modelToDoc } from './model-adapter';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const run = (text: string) => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;

async function open(paragraphXml: string) {
	const zip = new JSZip();
	zip.file(
		'word/document.xml',
		`<w:document xmlns:w="${W}"><w:body>${paragraphXml}<w:sectPr/></w:body></w:document>`,
	);
	const loaded = await loadDocx(await zip.generateAsync({ type: 'uint8array' }));
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
});
