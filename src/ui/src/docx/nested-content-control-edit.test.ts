// Editing a paragraph with nested inline content controls keeps the outer controls on its runs, so
// Print Layout still draws every tag after the edit (issue 374).
import JSZip from 'jszip';
import { EditorState } from 'prosemirror-state';
import { expect, it } from 'vitest';
import { loadDocx } from 'ooxml-core/docx';
import { createFakeMeasurer, layoutDocumentModel } from 'ooxml-core/docx/layout';
import { docToModel, modelToDoc } from './model-adapter';
import { schema } from './schema';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const W15 = 'http://schemas.microsoft.com/office/word/2012/wordml';
const sdt = (name: string, id: number, content: string) =>
	`<w:sdt><w:sdtPr><w:alias w:val="${name}"/><w:id w:val="${id}"/><w15:appearance w15:val="tags"/></w:sdtPr><w:sdtContent>${content}</w:sdtContent></w:sdt>`;
const run = (text: string) => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;

it('keeps the outer content control tags after an edit', async () => {
	const zip = new JSZip();
	zip.file(
		'word/document.xml',
		`<w:document xmlns:w="${W}" xmlns:w15="${W15}"><w:body><w:p>${run('Lorem ')}${sdt('Outer', 2, sdt('Inner', 1, run('ipsum')) + run(' sit'))}</w:p><w:sectPr/></w:body></w:document>`,
	);
	const { model } = await loadDocx(await zip.generateAsync({ type: 'uint8array' }));
	const state = EditorState.create({ schema, doc: modelToDoc(model) });
	const edited = state.apply(state.tr.insertText('X', 3));
	const result = docToModel(edited.doc, model);
	const { pages } = layoutDocumentModel(result, createFakeMeasurer());
	const block = pages[0]!.columns[0]!.blocks[0]!;
	if (block.kind !== 'paragraph') throw new Error('Expected paragraph');
	const tags = block.lines.flatMap((line) =>
		line.fragments.flatMap((fragment) =>
			fragment.contentControlTag ? [`${fragment.contentControlTag}:${fragment.text}`] : [],
		),
	);
	expect(tags).toEqual(['start:Outer', 'start:Inner', 'end:', 'end:']);
});
