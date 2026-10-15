import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { editVsdx, type VisioEdit } from './edit';
import { parseVsdx } from './parser';
import { fixture, shape } from './test-fixtures';
import { visioPointGlue } from './edit-connector-glue';

const PAGE = 'visio/pages/page1.xml';
const c = (name: string, value: string | number, formula = '', unit = '') =>
	`<Cell N="${name}" V="${value}"${unit ? ` U="${unit}"` : ''}${formula ? ` F="${formula}"` : ''}/>`;
const row = (index: number, type: string, cells: string) =>
	`<Row T="${type}" IX="${index}">${cells}</Row>`;

/** A 2-D master, one inch by three quarters, with a connection point on each side. */
const BOX =
	c('PinX', 2) +
	c('PinY', 2) +
	c('Width', 1) +
	c('Height', 0.75) +
	c('LocPinX', 0.5, 'Width*0.5') +
	c('LocPinY', 0.375, 'Height*0.5') +
	c('Angle', 0) +
	c('ObjType', 1) +
	`<Section N="Connection">${row(0, 'Connection', c('X', 0) + c('Y', 0.375, 'Height*0.5'))}${row(1, 'Connection', c('X', 1, 'Width') + c('Y', 0.375, 'Height*0.5'))}</Section>` +
	`<Section N="Geometry" IX="0">${row(1, 'MoveTo', c('X', 0, 'Width*0') + c('Y', 0, 'Height*0'))}${row(2, 'LineTo', c('X', 1, 'Width*1') + c('Y', 0, 'Height*0'))}${row(3, 'LineTo', c('X', 1, 'Width*1') + c('Y', 0.75, 'Height*1'))}${row(4, 'LineTo', c('X', 0, 'Width*0') + c('Y', 0.75, 'Height*1'))}${row(5, 'LineTo', c('X', 0, 'Geometry1.X1') + c('Y', 0, 'Geometry1.Y1'))}</Section>`;

/**
 * The box form of a line master, as the Connectors stencil's Angled connector is built (recorded
 * with scripts/record-visio-stencil-line.ps1): the box spans the ends, a flip flag turns it.
 */
const BOX_LINE = (pinX = 'GUARD((BeginX+EndX)/2)') =>
	c('PinX', 1.5, pinX) +
	c('PinY', 1.5, 'GUARD((BeginY+EndY)/2)') +
	c('Width', 1.2, 'GUARD(EndX-BeginX)') +
	c('Height', 1.6, 'GUARD(ABS(EndY-BeginY))') +
	c('LocPinX', 0.6, 'Width*0.5') +
	c('LocPinY', 0.8, 'Height*0.5') +
	c('Angle', 0, '', 'DEG') +
	c('FlipX', 0, 'GUARD(0)') +
	c('FlipY', 1, 'GUARD(BeginY>EndY)') +
	c('BeginX', 1) +
	c('BeginY', 2) +
	c('EndX', 2) +
	c('EndY', 1) +
	c('ObjType', 2) +
	`<Section N="Geometry" IX="0">${c('NoFill', 1)}${row(1, 'MoveTo', c('X', 0, 'Width*0') + c('Y', 0, 'Height*0'))}${row(2, 'LineTo', c('X', 1.2, 'Width*1') + c('Y', 1.6, 'Height*1'))}</Section>`;
/** The rotated form: the run's length, and its angle from the ends. */
const ROTATED_LINE =
	c('PinX', 1.5, '(BeginX+EndX)/2') +
	c('PinY', 1.5, '(BeginY+EndY)/2') +
	c('Width', 1.414213562373095, 'SQRT((EndX-BeginX)^2+(EndY-BeginY)^2)') +
	c('Height', 0) +
	c('LocPinX', 0.7071067811865475, 'Width*0.5') +
	c('LocPinY', 0, 'Height*0.5') +
	c('Angle', -Math.PI / 4, 'ATAN2(EndY-BeginY,EndX-BeginX)') +
	c('FlipX', 0) +
	c('FlipY', 0) +
	c('BeginX', 1) +
	c('BeginY', 2) +
	c('EndX', 2) +
	c('EndY', 1) +
	c('ObjType', 2) +
	`<Section N="Geometry" IX="0">${c('NoFill', 1)}${row(1, 'MoveTo', c('X', 0, 'Width*0') + c('Y', 0, 'Height*0.5'))}${row(2, 'LineTo', c('X', 1.414213562373095, 'Width*1') + c('Y', 0, 'Height*0.5'))}</Section>`;

const box = (id: string, x: number, y: number) =>
	shape(id, c('PinX', x) + c('PinY', y), 'Type="Shape" Master="2"');
/** The box line as Visio saves it from (2.5,6) to (5.5,4): glued to the right of 1 and the left of 2. */
const GLUED =
	c('PinX', 4, 'Inh', 'IN') +
	c('PinY', 5, 'Inh', 'IN') +
	c('Width', 3, 'Inh', 'IN') +
	c('Height', 2, 'Inh', 'IN') +
	c('LocPinX', 1.5, 'Inh', 'IN') +
	c('LocPinY', 1, 'Inh', 'IN') +
	c('BeginX', 2.5, visioPointGlue('1', 1), 'IN') +
	c('BeginY', 6, visioPointGlue('1', 1), 'IN') +
	c('EndX', 5.5, visioPointGlue('2', 0), 'IN') +
	c('EndY', 4, visioPointGlue('2', 0), 'IN') +
	`<Section N="Geometry" IX="0">${row(2, 'LineTo', c('X', 3, 'Inh', 'IN') + c('Y', 2, 'Inh', 'IN'))}</Section>`;
const CONNECTS =
	'<Connects><Connect FromSheet="3" FromCell="EndX" FromPart="12" ToSheet="2" ToCell="Connections.X1" ToPart="100"/><Connect FromSheet="3" FromCell="BeginX" FromPart="9" ToSheet="1" ToCell="Connections.X2" ToPart="101"/></Connects>';
const source = (
	options: { master?: string; connector?: string; connects?: string; page?: string } = {},
) =>
	fixture({
		masters: [
			{ id: '2', shapes: shape('5', BOX, 'Type="Shape"') },
			{ id: '4', shapes: shape('5', options.master ?? BOX_LINE(), 'Type="Shape"') },
		],
		pages: [
			{
				id: '0',
				contents: `<Shapes>${options.page ?? box('1', 2, 6) + box('2', 6, 4)}${shape('3', options.connector ?? GLUED, 'Type="Shape" Master="4"')}</Shapes>${options.connects ?? CONNECTS}`,
			},
		],
	});
const page = async (bytes: Uint8Array) =>
	(await JSZip.loadAsync(bytes)).file(PAGE)!.async('string');
const cells = async (bytes: Uint8Array, id = '3') => {
	const text = (await page(bytes)).replace(/'/g, '"');
	const body = new RegExp(`<Shape ID="${id}".*?</Shape>`, 's').exec(text)![0];
	return new Map(
		[...body.matchAll(/<Cell N="(\w+)" V="([^"]*)"(?: U="\w+")?(?: F="([^"]*)")?\/>/g)].map(
			(match) => [match[1]!, { v: match[2]!, f: match[3] }] as const,
		),
	);
};
const number = (map: Awaited<ReturnType<typeof cells>>, name: string) => Number(map.get(name)?.v);
const move = (shapeId: string, x: number, y: number): VisioEdit => ({
	type: 'move-shape',
	pageId: '0',
	shapeId,
	x,
	y,
});
const edited = async (bytes: Uint8Array, ...edits: VisioEdit[]) =>
	(await editVsdx(bytes, edits)).bytes;
const refused = async (bytes: Uint8Array, edits: VisioEdit[]) =>
	editVsdx(bytes, edits).then(
		() => undefined,
		(error: { code?: string; message: string }) => error,
	);

describe('a straight line master from a stencil (not the Dynamic connector)', () => {
	it('follows a moved shape, refreshing the inherited caches as Visio does', async () => {
		const bytes = await source();
		// Shape 2 moves up: its left connection point is now (5.5, 7).
		const moved = await cells(await edited(bytes, move('2', 6, 7)));
		expect([number(moved, 'BeginX'), number(moved, 'BeginY')]).toEqual([2.5, 6]);
		expect([number(moved, 'EndX'), number(moved, 'EndY')]).toEqual([5.5, 7]);
		expect(moved.get('EndX')!.f).toBe(visioPointGlue('2', 0));
		expect([number(moved, 'PinX'), number(moved, 'PinY')]).toEqual([4, 6.5]);
		expect([number(moved, 'Width'), number(moved, 'Height')]).toEqual([3, 1]);
		expect([number(moved, 'LocPinX'), number(moved, 'LocPinY')]).toEqual([1.5, 0.5]);
		// The ends now rise, so the box is no longer flipped: a refreshed local flag.
		expect(moved.get('FlipY')).toEqual({ v: '0', f: 'Inh' });
		expect(moved.get('PinX')!.f).toBe('Inh');
		expect(moved.get('Width')!.f).toBe('Inh');
		const again = await edited(await edited(bytes, move('2', 6, 7)), move('2', 6, 4));
		const back = await cells(again);
		// Back at the master's own flip, Visio drops the local flag.
		expect(back.has('FlipY')).toBe(false);
		expect([number(back, 'Width'), number(back, 'Height')]).toEqual([3, 2]);
		const model = await parseVsdx(again);
		expect(model.pages[0]!.connectors).toHaveLength(2);
	});

	it('follows a resized shape and a moved end shape on the other side', async () => {
		const bytes = await source();
		const wider = await edited(bytes, {
			type: 'resize-shape',
			pageId: '0',
			shapeId: '1',
			width: 2,
			height: 0.75,
			anchor: { x: 0, y: 0.5 },
		});
		const result = await cells(wider);
		// Anchored on its left side (x=1.5), shape 1 now ends at x=3.5, where its connection point is.
		expect([number(result, 'BeginX'), number(result, 'BeginY')]).toEqual([3.5, 6]);
		expect(number(result, 'Width')).toBeCloseTo(2, 12);
		expect(number(result, 'PinX')).toBeCloseTo(4.5, 12);
		const first = await cells(await edited(bytes, move('1', 1, 8)));
		expect([number(first, 'BeginX'), number(first, 'BeginY')]).toEqual([1.5, 8]);
		expect([number(first, 'Width'), number(first, 'Height')]).toEqual([4, 4]);
	});

	it('lays out the rotated form with an angle and the length of the run', async () => {
		const rotated = `${c('PinX', 4, 'Inh', 'IN')}${c('PinY', 5, 'Inh', 'IN')}${c('Width', Math.hypot(3, 2), 'Inh', 'IN')}${c('LocPinX', Math.hypot(3, 2) / 2, 'Inh', 'IN')}${c('Angle', Math.atan2(-2, 3), 'Inh')}${GLUED.slice(GLUED.indexOf('<Cell N="BeginX"'), GLUED.indexOf('<Section'))}`;
		const bytes = await source({ master: ROTATED_LINE, connector: rotated });
		const moved = await cells(await edited(bytes, move('2', 6, 7)));
		expect([number(moved, 'EndX'), number(moved, 'EndY')]).toEqual([5.5, 7]);
		expect(number(moved, 'Angle')).toBeCloseTo(Math.atan2(1, 3), 12);
		expect(number(moved, 'Width')).toBeCloseTo(Math.hypot(3, 1), 12);
		expect(number(moved, 'LocPinX')).toBeCloseTo(Math.hypot(3, 1) / 2, 12);
		expect([number(moved, 'PinX'), number(moved, 'PinY')]).toEqual([4, 6.5]);
		expect(moved.get('Angle')!.f).toBe('Inh');
	});

	it('moves its own end and unglues it, keeping the other end glued', async () => {
		const bytes = await source();
		const dragged = await edited(bytes, {
			type: 'move-line-endpoint',
			pageId: '0',
			shapeId: '3',
			endpoint: 'end',
			x: 7,
			y: 8,
		});
		const result = await cells(dragged);
		expect([number(result, 'EndX'), number(result, 'EndY')]).toEqual([7, 8]);
		expect(result.get('EndX')!.f).toBeUndefined();
		expect(result.get('BeginX')!.f).toBe(visioPointGlue('1', 1));
		expect([number(result, 'Width'), number(result, 'Height')]).toEqual([4.5, 2]);
		const text = await page(dragged);
		expect(text.match(/<Connect /g)).toHaveLength(1);
	});

	it('leaves a straight line straight and refuses to restyle it', async () => {
		const bytes = await source();
		const same = await editVsdx(bytes, [
			{ type: 'set-connector-route', pageId: '0', shapeId: '3', route: 'straight' },
		]);
		expect(await cells(same.bytes)).toEqual(await cells(bytes));
		const error = await refused(bytes, [
			{ type: 'set-connector-route', pageId: '0', shapeId: '3', route: 'curved' },
		]);
		expect(error?.code).toBe('UNSUPPORTED_GEOMETRY_EDIT');
		expect(error?.message).toMatch(/straight line from a stencil/);
	});

	it('refuses, leaving the file as it was, a master whose ends do not move its drawing', async () => {
		// The pin is a constant: the ends could move while the drawing stayed where it is.
		const fixed = await source({ master: BOX_LINE().replace(' F="GUARD((BeginX+EndX)/2)"', '') });
		const error = await refused(fixed, [move('2', 6, 7)]);
		expect(error?.code).toBe('UNSUPPORTED_GEOMETRY_EDIT');
		expect(error?.message).toMatch(/cannot be rerouted/);
		// A drawing whose path does not span the ends (the box is not stretched by the geometry).
		const wrong = await source({
			master: BOX_LINE().replace('GUARD(EndX-BeginX)', 'GUARD((EndX-BeginX)*2)'),
		});
		const mismatch = await refused(wrong, [move('2', 6, 7)]);
		expect(mismatch?.code).toMatch(/^UNSUPPORTED_(GEOMETRY|INSTANCE)_EDIT$/);
	});
});
