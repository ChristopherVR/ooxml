import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { editVsdx, type VisioEdit } from './edit';
import { parseVsdx } from './parser';
import { fixture, shape } from './test-fixtures';

const PAGE = 'visio/pages/page1.xml';
const MASTER = 'visio/masters/master1.xml';
const c = (name: string, value: string | number, formula = '', unit = '') =>
	`<Cell N="${name}" V="${value}"${unit ? ` U="${unit}"` : ''}${formula ? ` F="${formula}"` : ''}/>`;
const len = (name: string, value: number, formula = '') => c(name, value, formula, 'IN');
const row = (index: number, type: string, cells: string) =>
	`<Row T="${type}" IX="${index}">${cells}</Row>`;

/** A Parallelogram-like master: one horizontal handle whose X drives two corners and a connection point. */
const transform =
	len('PinX', 2) +
	len('PinY', 2) +
	len('Width', 2) +
	len('Height', 1) +
	len('LocPinX', 1, 'Width*0.5') +
	len('LocPinY', 0.5, 'Height*0.5') +
	c('Angle', 0) +
	c('FlipX', 0) +
	c('FlipY', 0);
const controls = (x: string, xCon = 0, yCon = 1) =>
	`<Section N="Control"><Row N="Row_1">${len('X', 0.2, x)}${len('Y', 1, 'Height*1')}${c('XCon', xCon)}${c('YCon', yCon)}${c('CanGlue', 0)}${c('Prompt', 'Slant')}</Row></Section>`;
const body = (options: { x?: string; xCon?: number; yCon?: number; extraGeometry?: string } = {}) =>
	transform +
	controls(options.x ?? 'BOUND(Width*0.1,0,FALSE,0,Width)', options.xCon, options.yCon) +
	`<Section N="Connection">${row(0, 'Connection', len('X', 0.1, 'Controls.Row_1*0.5') + len('Y', 0.5, 'Height*0.5'))}</Section>` +
	`<Section N="Geometry" IX="0">${c('NoFill', 0)}${row(1, 'MoveTo', len('X', 0.2, 'Controls.Row_1') + len('Y', 0, 'Height*0'))}${row(2, 'LineTo', len('X', 2, 'Width*1') + len('Y', 0, 'Height*0'))}${row(3, 'LineTo', len('X', 1.8, 'Width-Controls.Row_1') + len('Y', 1, 'Height*1'))}${row(4, 'LineTo', len('X', 0, 'Width*0') + len('Y', 1, 'Height*1'))}${options.extraGeometry ?? ''}</Section>`;
const master = (options: Parameters<typeof body>[0] = {}) =>
	shape('6', body(options), 'Type="Shape"');
const instance = `<Shape ID="1" Type="Shape" Master="2">${c('PinX', 2)}${c('PinY', 6)}</Shape>`;
const source = (definition = master(), page = instance) =>
	fixture({
		masters: [{ id: '2', shapes: definition }],
		pages: [{ id: '0', contents: `<Shapes>${page}</Shapes>` }],
	});
const local = (options: Parameters<typeof body>[0] = {}) =>
	`<Shape ID="1" Type="Shape">${body(options)}</Shape>`;
const part = async (bytes: Uint8Array, path: string) =>
	(await JSZip.loadAsync(bytes)).file(path)!.async('string');
const drag = (changes: { x?: number; y?: number }, row = 'Row_1'): VisioEdit => ({
	type: 'set-control',
	pageId: '0',
	shapeId: '1',
	row,
	...changes,
});
const refusal = async (bytes: Uint8Array, edit: VisioEdit) =>
	(
		await editVsdx(bytes, [edit]).then(
			() => undefined,
			(caught: { code?: string }) => caught,
		)
	)?.code;

describe('control handles in the model', () => {
	it('reads the handle of a stencil instance from its master', async () => {
		const model = (await parseVsdx(await source())).pages[0]!.shapes[0]!;
		expect(model.controls).toEqual([
			{
				row: 'Row_1',
				name: 'Row_1',
				x: 0.2,
				y: 1,
				prompt: 'Slant',
				canGlue: false,
				drag: { x: true, y: false },
				inherited: true,
			},
		]);
	});

	it('explains a handle that cannot be dragged', async () => {
		const computed = await source(master({ x: 'Width*0.5+User.Offset' }));
		const [handle] = (await parseVsdx(computed)).pages[0]!.shapes[0]!.controls!;
		expect(handle!.drag).toEqual({ x: false, y: false });
		expect(handle!.locked).toMatch(/formula/);
		const fixed = await source(master({ xCon: 1 }));
		const [pinned] = (await parseVsdx(fixed)).pages[0]!.shapes[0]!.controls!;
		expect(pinned!.locked).toMatch(/fixed/);
	});
});

describe('dragging a control handle of a stencil instance', () => {
	it('stores a ratio inside the cell limits and refreshes what follows it', async () => {
		const bytes = await source();
		const saved = await editVsdx(bytes, [drag({ x: 0.6 })]);
		expect(saved.changedParts).toEqual([PAGE]);
		expect(await part(saved.bytes, MASTER)).toBe(await part(bytes, MASTER));
		const page = await part(saved.bytes, PAGE);
		expect(page).toContain(
			'<Section N="Control"><Row N="Row_1"><Cell N="X" V="0.6" F="BOUND(Width*0.3,0,FALSE,0,Width)"/></Row></Section>',
		);
		for (const cache of [
			'<Cell N="X" V="0.3" U="IN" F="Inh"/>',
			'<Row T="MoveTo" IX="1"><Cell N="X" V="0.6" U="IN" F="Inh"/></Row>',
			'<Row T="LineTo" IX="3"><Cell N="X" V="1.4" U="IN" F="Inh"/></Row>',
		])
			expect(page).toContain(cache);
		const model = (await parseVsdx(saved.bytes)).pages[0]!.shapes[0]!;
		expect(model.controls![0]!.x).toBe(0.6);
		expect(model.connectionPoints![0]!.x).toBe(0.3);
	});

	it('keeps the handle inside its BOUND limits, as the cell does', async () => {
		const saved = await editVsdx(await source(), [drag({ x: 5 })]);
		const page = await part(saved.bytes, PAGE);
		expect(page).toContain('<Cell N="X" V="2" F="BOUND(Width*2.5,0,FALSE,0,Width)"/>');
		expect((await parseVsdx(saved.bytes)).pages[0]!.shapes[0]!.controls![0]!.x).toBe(2);
	});

	it('refuses an axis the handle is fixed on, and a formula it cannot rewrite', async () => {
		expect(await refusal(await source(), drag({ y: 0.5 }))).toBe('EDIT_PROTECTED_CELL');
		expect(await refusal(await source(master({ xCon: 1 })), drag({ x: 0.5 }))).toBe(
			'EDIT_PROTECTED_CELL',
		);
		expect(
			await refusal(await source(master({ x: 'Width*0.5+Height*0.1' })), drag({ x: 0.5 })),
		).toBe('EDIT_UNSUPPORTED_DEPENDENCY');
		expect(await refusal(await source(), drag({ x: 0.5 }, 'Nope'))).toBe('EDIT_TARGET_NOT_FOUND');
	});

	it('refuses a dependent cell it cannot recompute and writes nothing', async () => {
		const unsupported = row(5, 'LineTo', len('X', 1, 'FOO(Controls.Row_1)') + len('Y', 0));
		const bytes = await source(master({ extraGeometry: unsupported }));
		expect(await refusal(bytes, drag({ x: 0.6 }))).toBe('EDIT_UNSUPPORTED_DEPENDENCY');
	});
});

describe('dragging a control handle of a local shape', () => {
	it('rewrites the cell and the cells computed from it', async () => {
		const bytes = await fixture({ pages: [{ id: '0', contents: `<Shapes>${local()}</Shapes>` }] });
		const saved = await editVsdx(bytes, [drag({ x: 0.6 })]);
		const page = await part(saved.bytes, PAGE);
		expect(page).toContain('<Cell N="X" V="0.6" U="IN" F="BOUND(Width*0.3,0,FALSE,0,Width)"/>');
		expect(page).toContain('<Cell N="X" V="0.6" U="IN" F="Controls.Row_1"/>');
		expect(page).toContain('<Cell N="X" V="1.4" U="IN" F="Width-Controls.Row_1"/>');
	});
});
