import { describe, expect, it } from 'vitest';
import { readFile } from 'node:fs/promises';
import JSZip from 'jszip';
import { loadDocx } from './index';
import { forEachContentControl, parseGlossaryTexts } from './glossary';
import type { ContentControl } from './model';
import { WORD_NS } from './xml';

const rels = 'http://schemas.openxmlformats.org/package/2006/relationships';
const office = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const entry = (name: string, body: string) =>
	`<w:docPart><w:docPartPr><w:name w:val="${name}"/><w:category><w:name w:val="General"/>` +
	`<w:gallery w:val="placeholder"/></w:category></w:docPartPr><w:docPartBody>${body}</w:docPartBody></w:docPart>`;
const glossary = (parts: string) =>
	`<w:glossaryDocument xmlns:w="${WORD_NS}"><w:docParts>${parts}</w:docParts></w:glossaryDocument>`;
const styled = (value: string) =>
	`<w:p><w:r><w:rPr><w:rStyle w:val="PlaceholderText"/></w:rPr><w:t xml:space="preserve">${value}</w:t></w:r></w:p>`;

async function withGlossary(body: string, parts: string) {
	const zip = new JSZip();
	zip.file(
		'word/document.xml',
		`<w:document xmlns:w="${WORD_NS}"><w:body>${body}<w:sectPr/></w:body></w:document>`,
	);
	zip.file(
		'word/_rels/document.xml.rels',
		`<Relationships xmlns="${rels}"><Relationship Id="rId9" Type="${office}/glossaryDocument" Target="glossary/document.xml"/></Relationships>`,
	);
	zip.file('word/glossary/document.xml', glossary(parts));
	return zip.generateAsync({ type: 'uint8array' });
}
const controls = (blocks: Parameters<typeof forEachContentControl>[0]) => {
	const found: ContentControl[] = [];
	forEachContentControl(blocks, (control) => found.push(control));
	return found;
};
const cellChoice = (id: string, part: string) =>
	`<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="3000"/></w:tblGrid><w:tr><w:sdt><w:sdtPr><w:id w:val="${id}"/>` +
	`<w:placeholder><w:docPart w:val="${part}"/></w:placeholder><w:showingPlcHdr/><w:dropDownList/></w:sdtPr>` +
	`<w:sdtContent><w:tc>${styled('Pick one')}</w:tc></w:sdtContent></w:sdt></w:tr></w:tbl>`;

describe('placeholder text from the glossary', () => {
	it('reads each entry by name, with its first paragraph as its text', () => {
		const texts = parseGlossaryTexts(
			glossary(
				entry('A1B2', `${styled('Choose a status.')}${styled('second')}`) +
					entry('C3', '<w:p><w:r><w:t>Name</w:t><w:tab/><w:t>here</w:t></w:r></w:p>'),
			),
		);
		expect([...texts]).toEqual([
			['A1B2', 'Choose a status.'],
			['C3', 'Name\there'],
		]);
	});

	it('gives every control the text of its glossary entry', async () => {
		const loaded = await loadDocx(
			await withGlossary(
				`<w:p><w:sdt><w:sdtPr><w:id w:val="1"/><w:placeholder><w:docPart w:val="P1"/></w:placeholder>` +
					`<w:showingPlcHdr/><w:text/></w:sdtPr><w:sdtContent><w:r><w:t>Old</w:t></w:r></w:sdtContent></w:sdt></w:p>` +
					cellChoice('2', 'P2') +
					cellChoice('3', 'missing'),
				entry('P1', styled('Enter the document number.')) + entry('P2', styled('Choose an item.')),
			),
		);
		expect(controls(loaded.model.blocks).map((control) => control.placeholderText)).toEqual([
			'Enter the document number.',
			'Choose an item.',
			undefined,
		]);
		expect(controls(loaded.model.blocks)[1]).toMatchObject({ level: 'cell', placeholder: 'P2' });
		// The glossary is read only: an unedited document saves as it was loaded.
		const bytes = await withGlossary(cellChoice('4', 'P2'), entry('P2', styled('Choose an item.')));
		const reopened = await loadDocx(bytes);
		expect(await reopened.save(reopened.model)).toEqual(bytes);
	});

	it('reads the placeholders Word writes for its built-in controls', async () => {
		const fixture = new URL(
			'./__fixtures__/content-controls/content-controls.docx',
			import.meta.url,
		);
		const { model } = await loadDocx(new Uint8Array(await readFile(fixture)));
		const texts = controls(model.blocks).map((control) => [control.kind, control.placeholderText]);
		expect(texts).toContainEqual(['text', 'Click or tap here to enter text.']);
		expect(texts).toContainEqual(['dropDownList', 'Choose an item.']);
		expect(texts).toContainEqual(['date', 'Click or tap to enter a date.']);
	});
});
