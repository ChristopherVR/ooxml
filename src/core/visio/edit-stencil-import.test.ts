import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { createVsdx } from './create-document';
import { editVsdx } from './edit';
import { parseVsdx } from './parser';
import { readVisioStencil } from './stencil-file';
import { visioFileMasterDropCommand } from './ui/master-drop';
import { cell, fixture, row, section } from './test-fixtures';

const box = section(
	'Geometry',
	row(1, 'RelMoveTo', cell('X', 0) + cell('Y', 0)) +
		row(2, 'RelLineTo', cell('X', 1) + cell('Y', 0)) +
		row(3, 'RelLineTo', cell('X', 1) + cell('Y', 1)) +
		row(4, 'RelLineTo', cell('X', 0) + cell('Y', 1)) +
		row(5, 'RelLineTo', cell('X', 0) + cell('Y', 0)),
);
const valve = (id: string) =>
	`<Shape ID="${id}" Type="Shape" LineStyle="3">` +
	cell('PinX', 0.5) +
	cell('PinY', 0.5) +
	cell('Width', 1) +
	cell('Height', 1) +
	cell('LocPinX', 0.5, 'Width*0.5') +
	cell('LocPinY', 0.5, 'Height*0.5') +
	section(
		'Connection',
		row(0, 'Connection', cell('X', 0.5, 'Width*0.5') + cell('Y', 1, 'Height*1')) +
			row(1, 'Connection', cell('X', 0.5, 'Width*0.5') + cell('Y', 0, 'Height*0')),
	) +
	section(
		'Property',
		`<Row N="Pressure">${cell('Label', 'Pressure')}${cell('Value', 'bar')}${cell('Type', 0)}</Row>`,
	) +
	box +
	'</Shape>';

/** A stencil package: a drawing package whose main part has the stencil content type. */
const stencil = () =>
	fixture({
		document:
			'<StyleSheets><StyleSheet ID="0" NameU="No Style"/><StyleSheet ID="3" NameU="Thick Line" LineStyle="0"><Cell N="LineWeight" V="0.05"/></StyleSheet></StyleSheets>',
		masters: [
			{
				id: '7',
				shapes: valve('1'),
				attributes: 'NameU="Valve" Name="Valve" UniqueID="{11111111-2222-3333-4444-555555555555}"',
				sheet: cell('PageWidth', 1) + cell('PageHeight', 1),
			},
		],
		edit(zip) {
			zip.file(
				'[Content_Types].xml',
				'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Override PartName="/visio/document.xml" ContentType="application/vnd.ms-visio.stencil.main+xml"/></Types>',
			);
		},
	});

describe('reading a stencil', () => {
	it('lists its masters with connection points and Shape Data on their shapes', async () => {
		const file = await readVisioStencil(await stencil());
		expect(file.masters.map((master) => master.name)).toEqual(['Valve']);
		const root = file.masters[0]!.shapes[0]!;
		expect(root.connectionPoints).toHaveLength(2);
		expect(root.shapeData?.map((item) => item.label)).toEqual(['Pressure']);
	});

	it('refuses legacy binary stencils and macro-enabled packages', async () => {
		await expect(
			readVisioStencil(new Uint8Array([0xd0, 0xcf, 0x11, 0xe0, 0, 0, 0, 0])),
		).rejects.toThrow(/\.vss/);
		const macro = await fixture({
			masters: [{ id: '1', shapes: valve('1') }],
			edit(zip) {
				zip.file(
					'[Content_Types].xml',
					'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Override PartName="/visio/document.xml" ContentType="application/vnd.ms-visio.stencil.macroEnabled.main+xml"/></Types>',
				);
			},
		});
		await expect(readVisioStencil(macro)).rejects.toThrow(/not supported/);
	});

	it('is not accepted as a drawing', async () => {
		await expect(parseVsdx(await stencil())).rejects.toThrow();
	});
});

describe('dropping a master of an opened stencil', () => {
	it('copies the master into the drawing, and the saved instance resolves it', async () => {
		const bytes = await stencil();
		const master = (await readVisioStencil(bytes)).masters[0]!;
		const blank = await createVsdx();
		const page = (await parseVsdx(blank)).pages[0]!;
		const saved = await editVsdx(blank, [
			visioFileMasterDropCommand(page, master, bytes, { x: 2, y: 2 }),
		]);
		const reopened = await parseVsdx(saved.bytes);
		expect(reopened.masters!.map((item) => item.name)).toEqual(['Valve']);
		const dropped = reopened.pages[0]!.shapes[0]!;
		expect(dropped.masterId).toBe(reopened.masters![0]!.id);
		// The instance reaches the master's connection points and Shape Data through it.
		expect(dropped.connectionPoints).toHaveLength(2);
		expect(dropped.shapeData?.map((item) => item.label)).toEqual(['Pressure']);
		const zip = await JSZip.loadAsync(saved.bytes);
		expect(await zip.file('visio/masters/masters.xml')!.async('string')).toContain(
			'UniqueID="{11111111-2222-3333-4444-555555555555}"',
		);
		// The stencil's style came with the master.
		const document = await zip.file('visio/document.xml')!.async('string');
		expect(document).toContain('NameU="Thick Line"');
		expect(await zip.file('visio/masters/master1.xml')!.async('string')).toMatch(
			/LineStyle="(?!3")\d+"/,
		);
	});

	it('reuses the master when the same one is dropped again', async () => {
		const bytes = await stencil();
		const master = (await readVisioStencil(bytes)).masters[0]!;
		const blank = await createVsdx();
		const once = await editVsdx(blank, [
			visioFileMasterDropCommand((await parseVsdx(blank)).pages[0]!, master, bytes),
		]);
		const page = (await parseVsdx(once.bytes)).pages[0]!;
		const twice = await editVsdx(once.bytes, [
			visioFileMasterDropCommand(page, master, bytes, { x: 4, y: 4 }),
		]);
		const model = await parseVsdx(twice.bytes);
		expect(model.masters).toHaveLength(1);
		expect(model.pages[0]!.shapes.map((shape) => shape.masterId)).toEqual(['1', '1']);
	});

	it('refuses a master that owns pictures', async () => {
		const bytes = await fixture({
			masters: [{ id: '1', shapes: valve('1') }],
			edit(zip) {
				zip.file(
					'[Content_Types].xml',
					'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Override PartName="/visio/document.xml" ContentType="application/vnd.ms-visio.stencil.main+xml"/></Types>',
				);
				zip.file(
					'visio/masters/_rels/master1.xml.rels',
					'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.microsoft.com/visio/2010/relationships/image" Target="../media/a.png"/></Relationships>',
				);
				zip.file('visio/media/a.png', new Uint8Array([1]));
			},
		});
		const master = (await readVisioStencil(bytes)).masters[0]!;
		const blank = await createVsdx();
		await expect(
			editVsdx(blank, [
				visioFileMasterDropCommand((await parseVsdx(blank)).pages[0]!, master, bytes),
			]),
		).rejects.toThrow(/pictures/);
	});
});
