import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { loadDocx } from './index';
import { selectListItem } from './content-control-edit';
import type { Paragraph, Table } from './model';
import { WORD_NS } from './xml';

const hand = async (body: string) => {
	const zip = new JSZip();
	zip.file(
		'word/document.xml',
		`<w:document xmlns:w="${WORD_NS}"><w:body>${body}<w:sectPr/></w:body></w:document>`,
	);
	return loadDocx(await zip.generateAsync({ type: 'uint8array' }));
};
const documentXml = async (bytes: Uint8Array) =>
	(await (await JSZip.loadAsync(bytes)).file('word/document.xml')!.async('string')) as string;
const properties = (id: string, type = '') =>
	`<w:sdtPr><w:id w:val="${id}"/><w:temporary/><w:showingPlcHdr/>${type}</w:sdtPr>`;
const placeholder = (value: string) =>
	`<w:r><w:rPr><w:rStyle w:val="PlaceholderText"/></w:rPr><w:t>${value}</w:t></w:r>`;
const list =
	'<w:dropDownList><w:listItem w:displayText="Primary" w:value="primary"/></w:dropDownList>';

describe('temporary content controls on save', () => {
	it('keeps the picked entry of a temporary cell drop-down and removes the control', async () => {
		const loaded = await hand(
			`<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="3000"/></w:tblGrid><w:tr>` +
				`<w:sdt>${properties('51', list)}<w:sdtContent><w:tc><w:p>${placeholder('Choose an item.')}</w:p></w:tc></w:sdtContent></w:sdt>` +
				`</w:tr></w:tbl>`,
		);
		const cell = (loaded.model.blocks[0] as Table).rows[0]![0]!;
		const change = selectListItem(cell.contentControl!, 0)!;
		expect(change.removed).toBe(true);
		delete cell.contentControl;
		cell.paragraphs[0]!.runs = [{ text: change.text }];
		const xml = await documentXml(await loaded.save(loaded.model));
		expect(xml).not.toContain('<w:sdt');
		expect(xml).toContain('<w:tr><w:tc><w:p><w:r><w:t>Primary</w:t></w:r></w:p></w:tc></w:tr>');
	});

	it('removes a temporary block control and keeps its paragraphs where they were', async () => {
		const loaded = await hand(
			`<w:p><w:r><w:t>Before</w:t></w:r></w:p>` +
				`<w:sdt>${properties('52')}<w:sdtContent><w:p>${placeholder('Type here.')}</w:p><w:p><w:r><w:t>Second</w:t></w:r></w:p></w:sdtContent></w:sdt>` +
				`<w:p><w:r><w:t>After</w:t></w:r></w:p>`,
		);
		const [, first, second] = loaded.model.blocks as Paragraph[];
		delete first!.contentControl;
		delete second!.contentControl;
		first!.runs = [{ text: 'Typed' }];
		const xml = await documentXml(await loaded.save(loaded.model));
		expect(xml).not.toContain('<w:sdt');
		expect(xml).toMatch(/Before<\/w:t>.*Typed<\/w:t>.*Second<\/w:t>.*After<\/w:t>/);
	});

	it('removes a temporary block control whose edit also deleted one of its paragraphs', async () => {
		const loaded = await hand(
			`<w:sdt>${properties('58')}<w:sdtContent><w:p>${placeholder('Type here.')}</w:p><w:p><w:r><w:t>Second</w:t></w:r></w:p></w:sdtContent></w:sdt>` +
				`<w:p><w:r><w:t>After</w:t></w:r></w:p>`,
		);
		const [first, , after] = loaded.model.blocks as Paragraph[];
		delete first!.contentControl;
		first!.runs = [{ text: 'TypedSecond' }];
		loaded.model.blocks = [first!, after!];
		const xml = await documentXml(await loaded.save(loaded.model));
		expect(xml).not.toContain('<w:sdt');
		expect(xml).toMatch(/TypedSecond<\/w:t>.*After<\/w:t>/);
	});

	it('keeps a control whose temporary setting is turned off', async () => {
		const loaded = await hand(
			`<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="3000"/></w:tblGrid><w:tr>` +
				`<w:sdt><w:sdtPr><w:id w:val="57"/><w:temporary w:val="0"/><w:showingPlcHdr/>${list}</w:sdtPr>` +
				`<w:sdtContent><w:tc><w:p>${placeholder('Choose an item.')}</w:p></w:tc></w:sdtContent></w:sdt>` +
				`</w:tr></w:tbl>`,
		);
		const cell = (loaded.model.blocks[0] as Table).rows[0]![0]!;
		expect(cell.contentControl!.temporary).toBeUndefined();
		const change = selectListItem(cell.contentControl!, 0)!;
		expect(change.removed).toBeUndefined();
		cell.contentControl = change.control;
		cell.paragraphs[0]!.runs = [{ text: change.text }];
		const xml = await documentXml(await loaded.save(loaded.model));
		expect(xml).toContain('<w:sdt><w:sdtPr><w:id w:val="57"/><w:temporary w:val="0"/>');
		expect(xml).toContain('<w:t>Primary</w:t>');
	});

	it('keeps a block control that still holds one of its paragraphs', async () => {
		const loaded = await hand(
			`<w:sdt>${properties('53')}<w:sdtContent><w:p><w:r><w:t>One</w:t></w:r></w:p><w:p><w:r><w:t>Two</w:t></w:r></w:p></w:sdtContent></w:sdt>`,
		);
		const [first] = loaded.model.blocks as Paragraph[];
		delete first!.contentControl;
		first!.runs = [{ text: 'One!' }];
		const xml = await documentXml(await loaded.save(loaded.model));
		expect(xml).toContain('<w:id w:val="53"/>');
		expect(xml).toContain('One!');
	});

	it('saves the temporary flag the model gives a control', async () => {
		const loaded = await hand(
			`<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="3000"/></w:tblGrid><w:tr>` +
				`<w:sdt>${properties('55', list)}<w:sdtContent><w:tc><w:p>${placeholder('Choose an item.')}</w:p></w:tc></w:sdtContent></w:sdt>` +
				`</w:tr></w:tbl><w:p><w:r><w:t xml:space="preserve">Name: </w:t></w:r></w:p>`,
		);
		const cell = (loaded.model.blocks[0] as Table).rows[0]![0]!;
		delete cell.contentControl!.temporary;
		const paragraph = loaded.model.blocks[1] as Paragraph;
		paragraph.runs.push({
			text: 'Type here.',
			contentControl: { id: '56', level: 'inline', kind: 'text', temporary: true },
		});
		const xml = await documentXml(await loaded.save(loaded.model));
		expect(xml).toContain('<w:id w:val="55"/><w:showingPlcHdr/>');
		expect(xml).toContain('<w:sdtPr><w:id w:val="56"/><w:temporary/><w:text/></w:sdtPr>');
	});

	it('removes a temporary inline control and keeps its text', async () => {
		const loaded = await hand(
			`<w:p><w:r><w:t xml:space="preserve">Name: </w:t></w:r><w:sdt>${properties('54', '<w:text/>')}<w:sdtContent>${placeholder('Type here.')}</w:sdtContent></w:sdt></w:p>`,
		);
		const paragraph = loaded.model.blocks[0] as Paragraph;
		paragraph.runs = [paragraph.runs[0]!, { text: 'A. Writer' }];
		const xml = await documentXml(await loaded.save(loaded.model));
		expect(xml).not.toContain('<w:sdt');
		expect(xml).toContain('<w:t>A. Writer</w:t>');
	});
});
