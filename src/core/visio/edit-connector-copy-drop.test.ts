import { expect, it } from 'vitest';
import captured from './__fixtures__/native-connector-copy-drop.json';
import { editVsdx } from './edit';
import { cells, numeric } from './edit-geometry-cells';
import { pageShapes } from './edit-instance-shape';
import { effectiveCells, effectiveRows, registerStencilShape } from './edit-stencil-connector';
import { masterTemplate } from './edit-text-scope';
import { visioLineJumpPaths } from './line-jumps';
import { VisioPackage } from './package';
import { parseVsdx } from './parser';
import { attribute } from './sheet';
import { cell, fixture, row, section, shape } from './test-fixtures';
import { visioDuplicateCommand } from './ui/shape-duplicate';
import type { VisioEdit } from './edit';

interface NativeShape {
	id: string;
	type: number;
	cells: Readonly<Record<string, { value: number; formula: string }>>;
	geometry: readonly {
		rows: readonly { type: number; values: readonly number[]; formulas?: readonly string[] }[];
	}[];
	glue: number;
	children: readonly NativeShape[];
}
const transformCells = [
	'BeginX',
	'BeginY',
	'EndX',
	'EndY',
	'PinX',
	'PinY',
	'Width',
	'Height',
	'LocPinX',
	'LocPinY',
	'Angle',
	'FlipX',
	'FlipY',
];
const kinds = new Map([
	[138, 'MoveTo'],
	[139, 'LineTo'],
	[140, 'ArcTo'],
]);
const constant = /^\s*-?[\d.]+(?:\s+(?:in|deg))?$/i;

/** Reconstruct only the observed ShapeSheet cells in an original package envelope. */
function shapeXml(
	native: NativeShape,
	{ master, dynamic = false }: { master?: string; dynamic?: boolean } = {},
): string {
	const attrs = `Type="${native.type === 2 ? 'Group' : 'Shape'}"${master ? ` Master="${master}"` : ''}`;
	const values = Object.entries(native.cells)
		.map(([name, item]) => {
			let formula = constant.test(item.formula) ? '' : item.formula;
			if (dynamic && name === 'ShapeRouteStyle' && item.value === 0) return '';
			if (dynamic && master) {
				if (['Angle', 'FlipX', 'FlipY'].includes(name)) return '';
				if (['PinX', 'PinY', 'LocPinX', 'LocPinY'].includes(name)) formula = 'Inh';
				if (formula === 'GUARD(0.25 in)') formula = 'GUARD(0.25DL)';
				if (name.endsWith('Trigger') && item.value === 1)
					formula = `_XFTRIGGER(Sheet.${native.id}!EventXFMod)`;
			}
			return cell(
				name,
				item.value,
				formula.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/"/g, '&quot;'),
			);
		})
		.join('');
	const geometry = native.geometry
		.map(({ rows }, index) =>
			section(
				'Geometry',
				(native.cells.BeginX ? cell('NoFill', 1) : '') +
					rows
						.map((item, at) =>
							row(
								at + 1,
								kinds.get(item.type)!,
								item.values
									.map((value, column) => {
										const formula = item.formulas?.[column] ?? '';
										return cell(
											['X', 'Y', 'A'][column]!,
											value,
											constant.test(formula) ? '' : formula,
										);
									})
									.join(''),
							),
						)
						.join(''),
				index,
			),
		)
		.join('');
	const members = native.children.length
		? `<Shapes>${native.children.map((child) => shapeXml(child)).join('')}</Shapes>`
		: '';
	return shape(native.id, values + cell('EventXFMod', 0) + geometry + members, attrs);
}

async function reopened(bytes: Uint8Array, id: string): Promise<Element> {
	const pkg = await VisioPackage.open(bytes);
	const root = await pkg.readXml('visio/pages/page1.xml');
	const result = pageShapes(root).find((el) => attribute(el, 'ID') === id)!;
	const master = attribute(result, 'Master');
	if (master) registerStencilShape(result, await masterTemplate(pkg)(master));
	return result;
}

function matchesCells(actual: Map<string, Element>, reference: NativeShape) {
	for (const name of transformCells)
		if (reference.cells[name])
			expect(numeric(actual.get(name)), name).toBeCloseTo(reference.cells[name].value, 10);
}

function matchesRoute(actual: Element, reference: NativeShape) {
	matchesCells(effectiveCells(actual), reference);
	const rows = [...effectiveRows(actual, 'Geometry')].sort(([a], [b]) => a - b);
	expect(rows).toHaveLength(reference.geometry[0]!.rows.length);
	for (const [[, values], native] of rows.map(
		(entry, at) => [entry, reference.geometry[0]!.rows[at]!] as const,
	))
		for (const [column, value] of native.values.entries())
			expect(numeric(values.get(['X', 'Y', 'A'][column]!))).toBeCloseTo(value, 5);
}

it.each(captured.masters)(
	'drops and duplicates $name with the saved Visio transform and geometry',
	async (native) => {
		const masterId = '2';
		const stencil = await fixture({
			masters: [
				{ id: masterId, shapes: shapeXml(native.master), attributes: `NameU="${native.name}"` },
			],
		});
		const initial = await fixture({ pages: [{ id: '0', contents: '' }] });
		const drop = await editVsdx(initial, [
			{
				type: 'drop-stencil-master',
				pageId: '0',
				shapeId: '1',
				master: masterId,
				stencil,
				x: 4,
				y: 5,
			},
		]);
		const dropped = await reopened(drop.bytes, '1');
		matchesCells(effectiveCells(dropped), native.drop);
		const page = (await parseVsdx(drop.bytes)).pages[0]!;
		const command = visioDuplicateCommand(page, ['1']);
		expect(command).toBeDefined();
		const copied = await editVsdx(drop.bytes, [command!]);
		const copyId = command!.copies[0]!.newShapeId;
		const reopenedCopy = await reopened(copied.bytes, copyId);
		matchesCells(effectiveCells(reopenedCopy), native.copy);
		const model = (await parseVsdx(copied.bytes)).pages[0]!;
		expect(model.connectors).toEqual([]);
		expect(model.shapes[1]?.masterId).toBe(model.shapes[0]?.masterId);
		expect(model.shapes[1]?.geometry).toEqual(model.shapes[0]?.geometry);
		expect(model.shapes[1]?.geometry).toHaveLength(native.copy.geometry.length);
		expect(model.shapes[1]?.children).toHaveLength(native.copy.children.length);
		if (native.copy.children.length) {
			const child = model.shapes[1]!.children[0]!;
			expect(child.id).not.toBe(model.shapes[0]!.children[0]!.id);
			expect(child.geometry).toEqual(model.shapes[0]!.children[0]!.geometry);
		}
	},
);

it('duplicates across a glued dynamic connector and saves the native jump arc', async () => {
	const { crossing, dynamic } = captured;
	const master = '2';
	const target = (id: string, pinY: number) =>
		shape(id, cell('PinX', 4) + cell('PinY', pinY) + cell('Width', 1) + cell('Height', 1));
	const connect = (end: string, id: string) =>
		`<Connect FromSheet="4" FromCell="${end}X" FromPart="${end === 'Begin' ? 9 : 12}" ToSheet="${id}" ToCell="PinX" ToPart="3"/>`;
	const source = await fixture({
		masters: [{ id: master, shapes: shapeXml(dynamic, { dynamic: true }) }],
		pages: [
			{
				id: '0',
				pageCells: cell('LineJumpCode', 1),
				contents:
					'<Shapes>' +
					shapeXml(crossing.source, { master, dynamic: true }) +
					target('2', 1.5) +
					target('3', 3.775) +
					shapeXml(crossing.otherBefore, { master, dynamic: true }) +
					'</Shapes><Connects>' +
					connect('Begin', '2') +
					connect('End', '3') +
					'</Connects>',
			},
		],
	});
	const page = (await parseVsdx(source)).pages[0]!;
	const edited = await editVsdx(source, [visioDuplicateCommand(page, ['1'])!]);
	const copy = await reopened(edited.bytes, '5');
	matchesRoute(copy, crossing.copy);
	matchesRoute(await reopened(edited.bytes, '4'), crossing.otherAfter);
	const saved = (await parseVsdx(edited.bytes)).pages[0]!;
	expect(saved.connectors).toHaveLength(crossing.otherAfter.glue);
	const drawn = saved.shapes.find((item) => item.id === '5')!;
	expect(drawn.geometry[0]!.path).toContain(' A ');
	expect(visioLineJumpPaths(saved).get(drawn)?.[0]).toContain(' A ');
	expect(attribute(cells(copy).get('BegTrigger'), 'F')).toBe('_XFTRIGGER(Sheet.5!EventXFMod)');
	const detached = await editVsdx(edited.bytes, [
		{
			type: 'duplicate-shapes',
			pageId: '0',
			copies: [{ shapeId: '4', newShapeId: '6' }],
			offsetX: 0,
			offsetY: 0,
		},
	]);
	const gluedCopy = await reopened(detached.bytes, '6');
	matchesRoute(gluedCopy, crossing.gluedCopy);
	expect(attribute(cells(gluedCopy).get('BeginX'), 'F')).toBe(
		crossing.gluedCopy.cells.BeginX.formula,
	);
	expect(numeric(effectiveCells(gluedCopy).get('ConFixedCode'))).toBe(0);
	expect((await parseVsdx(detached.bytes)).pages[0]?.connectors).toHaveLength(2);
});

it.each(captured.moves)(
	'reroutes a free connector as the obstacle moves on $axis',
	async (native) => {
		const master = '2';
		const connector = shapeXml(native.before, { master, dynamic: true });
		const source = await fixture({
			masters: [{ id: master, shapes: shapeXml(captured.dynamic, { dynamic: true }) }],
			pages: [
				{
					id: '0',
					contents: `<Shapes>${connector}${shapeXml(native.blocker)}</Shapes>`,
				},
			],
		});
		const move = (value: number): VisioEdit => ({
			type: 'move-shape',
			pageId: '0',
			shapeId: native.blocker.id,
			x: native.axis === 'X' ? value : 4,
			y: native.axis === 'Y' ? value : 4,
		});
		const blocked = await editVsdx(source, [move(4)]);
		matchesRoute(await reopened(blocked.bytes, native.before.id), native.blocked);
		const cleared = await editVsdx(blocked.bytes, [move(6)]);
		matchesRoute(await reopened(cleared.bytes, native.before.id), native.cleared);
		expect((await parseVsdx(cleared.bytes)).pages[0]?.connectors).toEqual([]);
	},
);
