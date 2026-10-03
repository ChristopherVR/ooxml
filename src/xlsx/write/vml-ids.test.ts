import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { createEditSession } from '../edit/index.js';
import { loadXlsx } from '../read/index.js';
import { saveXlsx } from './index.js';
import { VmlIdAllocator, stripNoteShapes, vmlIdmapBlocks, vmlShapeIds } from './vml-ids.js';

const MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PKG = 'http://schemas.openxmlformats.org/package/2006/relationships';
const VML_NS =
	'xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel"';

// Modelled on what Excel 16 writes for a sheet with a form-control check box and one note.
const CHECKBOX_VML =
	`<xml ${VML_NS}><o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="1"/></o:shapelayout>` +
	'<v:shapetype id="_x0000_t201" coordsize="21600,21600" o:spt="201" path="m,l,21600r21600,l21600,xe"><v:path o:connecttype="rect"/></v:shapetype>' +
	'<v:shape id="_x0000_s1025" type="#_x0000_t201" style="position:absolute;width:80pt;height:20pt"><v:textbox><div>Tick me</div></v:textbox>' +
	'<x:ClientData ObjectType="Checkbox"><x:Anchor>2, 8, 2, 2, 3, 72, 3, 13</x:Anchor><x:FmlaLink>$C$1</x:FmlaLink></x:ClientData></v:shape>' +
	'<v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe"><v:path gradientshapeok="t" o:connecttype="rect"/></v:shapetype>' +
	'<v:shape id="_x0000_s1026" type="#_x0000_t202" style="position:absolute;visibility:hidden"><v:textbox><div/></v:textbox>' +
	'<x:ClientData ObjectType="Note"><x:Anchor>5, 15, 3, 10, 7, 15, 7, 4</x:Anchor><x:Row>4</x:Row><x:Column>4</x:Column></x:ClientData></v:shape></xml>';

async function formControlWorkbook(): Promise<Uint8Array> {
	const zip = new JSZip();
	zip.file(
		'[Content_Types].xml',
		'<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/comments1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"/><Override PartName="/xl/ctrlProps/ctrlProp1.xml" ContentType="application/vnd.ms-excel.controlproperties+xml"/></Types>',
	);
	zip.file(
		'_rels/.rels',
		`<Relationships xmlns="${PKG}"><Relationship Id="rId1" Type="${REL}/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
	);
	zip.file(
		'xl/workbook.xml',
		`<workbook xmlns="${MAIN}" xmlns:r="${REL}"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>`,
	);
	zip.file(
		'xl/_rels/workbook.xml.rels',
		`<Relationships xmlns="${PKG}"><Relationship Id="rId1" Type="${REL}/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`,
	);
	zip.file(
		'xl/worksheets/sheet1.xml',
		`<worksheet xmlns="${MAIN}" xmlns:r="${REL}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>x</t></is></c></row></sheetData><legacyDrawing r:id="rId2"/><mc:AlternateContent><mc:Choice Requires="x14"><controls><mc:AlternateContent><mc:Choice Requires="x14"><control shapeId="1025" r:id="rId3" name="Check Box 1"/></mc:Choice></mc:AlternateContent></controls></mc:Choice></mc:AlternateContent></worksheet>`,
	);
	zip.file(
		'xl/worksheets/_rels/sheet1.xml.rels',
		`<Relationships xmlns="${PKG}"><Relationship Id="rId2" Type="${REL}/vmlDrawing" Target="../drawings/vmlDrawing1.vml"/><Relationship Id="rId3" Type="${REL}/ctrlProp" Target="../ctrlProps/ctrlProp1.xml"/><Relationship Id="rId4" Type="${REL}/comments" Target="../comments1.xml"/></Relationships>`,
	);
	zip.file('xl/drawings/vmlDrawing1.vml', CHECKBOX_VML);
	zip.file(
		'xl/comments1.xml',
		`<comments xmlns="${MAIN}"><authors><author>Author</author></authors><commentList><comment ref="E5" authorId="0" shapeId="0"><text><r><t>note here</t></r></text></comment></commentList></comments>`,
	);
	zip.file(
		'xl/ctrlProps/ctrlProp1.xml',
		'<formControlPr xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" objectType="CheckBox" fmlaLink="$C$1"/>',
	);
	return zip.generateAsync({ type: 'uint8array' });
}

/** Every VML drawing of a saved package. */
async function vmlParts(bytes: Uint8Array): Promise<Map<string, string>> {
	const zip = await JSZip.loadAsync(bytes);
	const out = new Map<string, string>();
	for (const name of Object.keys(zip.files).filter((n) => n.endsWith('.vml')))
		out.set(name, (await zip.file(name)?.async('string')) ?? '');
	return out;
}

function expectNoCollisions(parts: Map<string, string>): void {
	const blocks = [...parts.values()].flatMap(vmlIdmapBlocks);
	expect(new Set(blocks).size).toBe(blocks.length);
	const ids = [...parts.values()].flatMap(vmlShapeIds);
	expect(new Set(ids).size).toBe(ids.length);
	for (const xml of parts.values()) {
		const own = vmlIdmapBlocks(xml);
		for (const id of vmlShapeIds(xml)) expect(own).toContain(Math.floor(id / 1024));
	}
}

describe('comment VML shape ids', () => {
	it('gives a new sheet comment a block no carried form-control drawing holds', async () => {
		const workbook = await loadXlsx(await formControlWorkbook());
		const session = createEditSession(workbook, { recalc: false });
		const added = session.addSheet('New', 0);
		session.setComment(added, { row: 0, col: 0 }, 'new note', 'me');
		const parts = await vmlParts(await saveXlsx(workbook));
		expect(parts.size).toBe(2);
		expect(parts.get('xl/drawings/vmlDrawing1.vml')).toContain('_x0000_s1025');
		expectNoCollisions(parts);
	});

	it('keeps the check box when a note on the same sheet is edited', async () => {
		const workbook = await loadXlsx(await formControlWorkbook());
		const comment = workbook.sheets[0]?.comments[0];
		if (comment) comment.text = 'edited';
		const parts = await vmlParts(await saveXlsx(workbook));
		const xml = [...parts.values()].join('');
		expect(xml.match(/ObjectType="Checkbox"/g)).toHaveLength(1);
		expect(xml.match(/ObjectType="Note"/g)).toHaveLength(1);
		expect(xml).toContain('id="_x0000_s1025"');
		expectNoCollisions(parts);
		const reloaded = await loadXlsx(await saveXlsx(workbook));
		expect(reloaded.sheets[0]?.comments.map((c) => c.text)).toEqual(['edited']);
	});

	it('keeps the check box drawing when every note is removed', async () => {
		const workbook = await loadXlsx(await formControlWorkbook());
		const sheet = workbook.sheets[0];
		if (sheet) sheet.comments = [];
		const bytes = await saveXlsx(workbook);
		const xml = [...(await vmlParts(bytes)).values()].join('');
		expect(xml).toContain('ObjectType="Checkbox"');
		expect(xml).not.toContain('ObjectType="Note"');
		const sheetXml = await (
			await JSZip.loadAsync(bytes)
		)
			.file('xl/worksheets/sheet1.xml')
			?.async('string');
		expect(sheetXml).toContain('<legacyDrawing');
	});
});

describe('VmlIdAllocator', () => {
	it('fills free slots of its own blocks first, then claims unused blocks', () => {
		const allocator = new VmlIdAllocator(undefined);
		const first = allocator.allocate(2, 1);
		expect(first).toEqual({ blocks: [1], ids: [1025, 1026] });
		const second = allocator.allocate(1, 1);
		expect(second).toEqual({ blocks: [2], ids: [2049] });
		const big = allocator.allocate(1100, 1);
		expect(big.blocks).toEqual([3, 4]);
		expect(new Set(big.ids).size).toBe(1100);
	});

	it('reads comma-separated idmap lists and strips only note shapes', () => {
		expect(vmlIdmapBlocks('<o:idmap v:ext="edit" data="3,4"/>')).toEqual([3, 4]);
		const stripped = stripNoteShapes(CHECKBOX_VML);
		expect(stripped.kept).toBe(1);
		expect(stripped.xml).not.toContain('ObjectType="Note"');
		expect(stripped.xml).toContain('ObjectType="Checkbox"');
	});
});
