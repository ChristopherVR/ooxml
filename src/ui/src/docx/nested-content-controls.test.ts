// @vitest-environment jsdom
// Editing inside nested block content controls (a repeating section's item, a group around a rich
// text control) through the editor's model adapter keeps every `w:sdt` around the paragraph when
// the document is saved: typing, and splitting a paragraph in its middle and at its start.
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { loadDocx } from 'ooxml-core/docx';
import { splitBlock } from 'prosemirror-commands';
import { EditorState, TextSelection } from 'prosemirror-state';
import { assignMissingParagraphIds, docToModel, modelToDoc } from './model-adapter';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const W15 = 'http://schemas.microsoft.com/office/word/2012/wordml';
const run = (text: string) => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;
const sdt = (id: number, properties: string, content: string) =>
	`<w:sdt><w:sdtPr><w:id w:val="${id}"/>${properties}</w:sdtPr><w:sdtContent>${content}</w:sdtContent></w:sdt>`;
const SECTION = sdt(
	1,
	`<w:tag w:val="sec"/><w15:repeatingSection xmlns:w15="${W15}"/>`,
	sdt(2, `<w15:repeatingSectionItem xmlns:w15="${W15}"/>`, `<w:p>${run('Item one')}</w:p>`),
);
const GROUP = sdt(
	3,
	'<w:tag w:val="grp"/><w:group/>',
	sdt(4, '<w:tag w:val="inner"/>', `<w:p>${run('Inner text')}</w:p>`),
);

async function load(body: string) {
	const zip = new JSZip();
	zip.file(
		'word/document.xml',
		`<w:document xmlns:w="${W}"><w:body>${body}<w:p>${run('After')}</w:p></w:body></w:document>`,
	);
	return loadDocx(await zip.generateAsync({ type: 'uint8array' }));
}

/** Edits the first paragraph in the editor (`at` from its start; split there when `split`). */
async function edit(body: string, at: number, split: boolean) {
	const loaded = await load(body);
	const doc = modelToDoc(loaded.model);
	let state = EditorState.create({ doc });
	state = state.apply(state.tr.setSelection(TextSelection.create(state.doc, 1 + at)));
	if (split) splitBlock(state, (tr) => (state = state.apply(tr)));
	state = state.apply(state.tr.insertText('typed'));
	const fix = assignMissingParagraphIds(state);
	if (fix) state = state.apply(fix);
	const model = docToModel(state.doc, loaded.model, doc);
	const bytes = await loaded.save(model);
	return (await (await JSZip.loadAsync(bytes)).file('word/document.xml')!.async('string')).replace(
		/\sxmlns:w15="[^"]*"/g,
		'',
	);
}

const ids = (xml: string) => [...xml.matchAll(/<w:id w:val="(\d+)"\/>/g)].map((match) => match[1]);

describe('editing inside nested block content controls', () => {
	it.each([
		['types inside', 4, false],
		['splits in the middle', 4, true],
		['splits at the start', 0, true],
	])('keeps a repeating section around its item when the editor %s', async (_name, at, split) => {
		const xml = await edit(SECTION, at, split);
		expect(xml.match(/<w15:repeatingSection\b/g)).toHaveLength(1);
		expect(xml.match(/<w15:repeatingSectionItem\b/g)).toHaveLength(1);
		expect(xml).toContain('<w:tag w:val="sec"/>');
		expect(ids(xml)).toEqual(['1', '2']);
		// Every paragraph of the item stays inside it, and "After" outside the section.
		const section = xml.slice(xml.indexOf('<w:sdt>'), xml.lastIndexOf('</w:sdt>'));
		expect(section).toContain('typed');
		expect(section).not.toContain('After');
	});

	it.each([
		['types inside', 4, false],
		['splits in the middle', 4, true],
		['splits at the start', 0, true],
	])('keeps a group around its rich text control when the editor %s', async (_name, at, split) => {
		const xml = await edit(GROUP, at, split);
		expect(xml).toContain('<w:group/>');
		expect(xml).toContain('<w:tag w:val="grp"/>');
		expect(ids(xml)).toEqual(['3', '4']);
		expect(xml.slice(0, xml.lastIndexOf('</w:sdt>'))).toContain('typed');
	});
});
