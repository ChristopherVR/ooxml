import { expect, it } from 'vitest';
import { editVsdx } from './edit';
import { VisioPackage } from './package';
import { attribute, children } from './sheet';
import { cell, fixture, rectangle, row, section, shape } from './test-fixtures';

const move = { type: 'move-shape' as const, pageId: '0', shapeId: '1', x: 4.25, y: 3 };
const endpoint = {
	type: 'move-line-endpoint' as const,
	pageId: '0',
	shapeId: '1',
	endpoint: 'end' as const,
	x: 4,
	y: 2,
};
function line(
	ex = 3,
	ey = 1.5,
	overrides: Record<string, [number, string?]> = {},
	geometry?: string,
) {
	const width = Math.hypot(ex - 1, ey - 1.5);
	const values: Record<string, [number, string?]> = {
		BeginX: [1],
		BeginY: [1.5],
		EndX: [ex],
		EndY: [ey],
		PinX: [(1 + ex) / 2, '(BeginX+EndX)/2'],
		PinY: [(1.5 + ey) / 2, '(BeginY+EndY)/2'],
		Width: [width, 'SQRT((EndX-BeginX)^2+(EndY-BeginY)^2)'],
		Height: [0],
		Angle: [Math.atan2(ey - 1.5, ex - 1), 'ATAN2(EndY-BeginY,EndX-BeginX)'],
		LocPinX: [width / 2, 'Width*0.5'],
		LocPinY: [0, 'Height*0.5'],
		...overrides,
	};
	return shape(
		'1',
		Object.entries(values)
			.map(([name, [value, formula]]) => cell(name, value, formula))
			.join('') +
			(geometry ??
				section(
					'Geometry',
					row(1, 'MoveTo', cell('X', 0, 'Width*0') + cell('Y', 0)) +
						row(2, 'LineTo', cell('X', width, 'Width*1') + cell('Y', 0)),
				)),
		'Type="Shape"',
	);
}
const source = (target = line(), extra = '', trailing = '') =>
	fixture({
		pages: [
			{
				id: '0',
				contents: `<Shapes>${target}${extra}</Shapes>${trailing}`,
			},
		],
	});
async function savedCells(bytes: Uint8Array, id = '1') {
	const pkg = await VisioPackage.open(bytes);
	const root = await pkg.readXml('visio/pages/page1.xml', 'PageContents');
	const target = children(children(root, 'Shapes')[0], 'Shape').find(
		(node) => attribute(node, 'ID') === id,
	);
	return new Map(children(target, 'Cell').map((node) => [attribute(node, 'N'), node]));
}

it.each([
	[3, 1.5],
	[2.732050807568877, 2.5],
	[1, 3.5],
	[-1, 1.5],
])('round-trips translated endpoints for end (%s,%s)', async (x, y) => {
	const bytes = await source(line(x!, y!));
	const before = await savedCells(bytes);
	const saved = await editVsdx(bytes, [move]);
	const after = await savedCells(saved.bytes);
	const dx = move.x - Number(attribute(before.get('PinX'), 'V'));
	const dy = move.y - Number(attribute(before.get('PinY'), 'V'));
	for (const [name, delta] of [
		['BeginX', dx],
		['EndX', dx],
		['BeginY', dy],
		['EndY', dy],
	] as const)
		expect(Number(attribute(after.get(name), 'V'))).toBeCloseTo(
			Number(attribute(before.get(name), 'V')) + delta,
			12,
		);
	for (const name of ['Width', 'Height', 'Angle', 'LocPinX', 'LocPinY']) {
		expect(Number(attribute(after.get(name), 'V'))).toBeCloseTo(
			Number(attribute(before.get(name), 'V')),
			12,
		);
		expect(attribute(after.get(name), 'F')).toBe(attribute(before.get(name), 'F'));
	}
	for (const [name, value] of [
		['PinX', move.x],
		['PinY', move.y],
	] as const) {
		expect(Number(attribute(after.get(name), 'V'))).toBeCloseTo(value, 12);
		expect(attribute(after.get(name), 'F')).toBe(attribute(before.get(name), 'F'));
	}
	const restored = await editVsdx(saved.bytes, [
		{
			...move,
			x: Number(attribute(before.get('PinX'), 'V')),
			y: Number(attribute(before.get('PinY'), 'V')),
		},
	]);
	const originalPosition = await savedCells(restored.bytes);
	expect(Number(attribute(originalPosition.get('BeginX'), 'V'))).toBeCloseTo(1, 12);
	expect(Number(attribute(originalPosition.get('BeginY'), 'V'))).toBeCloseTo(1.5, 12);
});

it('recalculates references to translated endpoints through the existing dependency closure', async () => {
	const other = shape(
		'2',
		cell('Width', 2) +
			cell('Height', 1) +
			cell('PinX', 3, 'Sheet.1!EndX') +
			cell('PinY', 4) +
			rectangle,
	);
	const saved = await editVsdx(await source(line(), other), [move]);
	const dependent = await savedCells(saved.bytes, '2');
	expect(Number(attribute(dependent.get('PinX'), 'V'))).toBe(5.25);
	expect(attribute(dependent.get('PinX'), 'F')).toBe('Sheet.1!EndX');
});

it.each([
	{ BeginX: [1, 'GUARD(1)'] },
	{ EndY: [1.5, 'Sheet.2!PinY'] },
	{ LockMoveX: [1] },
	{ LockMoveY: [1] },
	{ LockBegin: [1] },
	{ LockEnd: [1] },
	{ LockBegin: [0, '1'] },
	{ LockEnd: [0, 'Sheet.2!User.lock'] },
	{ PinX: [99, '(BeginX+EndX)/2'] },
	{ Width: [3, 'SQRT((EndX-BeginX)^2+(EndY-BeginY)^2)'] },
	{ Angle: [0, '(PinX-2in)/1in'] },
	{ Height: [0, 'PinX-2in'] },
	{ Width: [2, '2rad'] },
	{ Width: [2, '3in'] },
] satisfies Record<string, [number, string?]>[])(
	'rejects protected, stale or pose-changing line cells (%s) atomically',
	async (overrides) => {
		const bytes = await source(line(3, 1.5, overrides));
		const original = bytes.slice();
		await expect(editVsdx(bytes, [move])).rejects.toThrow();
		expect(bytes).toEqual(original);
	},
);

it('keeps an exact no-op unchanged even when endpoint constants are guarded', async () => {
	const bytes = await source(line(3, 1.5, { BeginX: [1, 'GUARD(1)'] }));
	const saved = await editVsdx(bytes, [{ ...move, x: 2, y: 1.5 }]);
	expect(saved.bytes).toEqual(bytes);
});

it('rejects glued lines, curved movement geometry, endpoint angle units and resizing', async () => {
	await expect(
		editVsdx(
			await source(
				line(),
				'',
				'<Connects><Connect FromSheet="1" ToSheet="2" FromCell="BeginX" ToCell="PinX"/></Connects>',
			),
			[move],
		),
	).rejects.toThrow('Glued connections');
	await expect(
		editVsdx(
			await source(
				line(
					3,
					1.5,
					{},
					section(
						'Geometry',
						row(1, 'MoveTo', cell('X', 0) + cell('Y', 0)) +
							row(2, 'ArcTo', cell('X', 2) + cell('Y', 0)),
					),
				),
			),
			[move],
		),
	).rejects.toThrow('straight-line');
	await expect(
		editVsdx(await source(line().replace('N="EndX"', 'N="EndX" U="DEG"')), [move]),
	).rejects.toThrow('length units');
	const bytes = await source();
	await expect(
		editVsdx(bytes, [{ type: 'resize-shape', pageId: '0', shapeId: '1', width: 3, height: 1 }]),
	).rejects.toThrow();
});

it.each(['LockBegin', 'LockEnd'])(
	'resolves inherited %s through shared style protection',
	async (lock) => {
		const bytes = await fixture({
			document: `<StyleSheets><StyleSheet ID="0">${cell(lock, 1)}</StyleSheet></StyleSheets>`,
			pages: [{ id: '0', contents: `<Shapes>${line()}</Shapes>` }],
		});
		await expect(editVsdx(bytes, [move])).rejects.toThrow('Inherited protection');
	},
);

it.each([
	[3, 1.5],
	[2.732050807568877, 2.5],
	[1, 3.5],
	[-1, 1.5],
])('resizes Width while preserving endpoint cells for end (%s,%s)', async (x, y) => {
	const bytes = await source(line(x!, y!));
	const resized = await editVsdx(bytes, [
		{ type: 'resize-shape', pageId: '0', shapeId: '1', width: 4, height: 0 },
	]);
	const before = await savedCells(bytes),
		after = await savedCells(resized.bytes);
	expect(Number(attribute(after.get('Width'), 'V'))).toBe(4);
	expect(attribute(after.get('Width'), 'F')).toBeUndefined();
	expect(Number(attribute(after.get('LocPinX'), 'V'))).toBe(2);
	for (const name of ['BeginX', 'BeginY', 'EndX', 'EndY', 'PinX', 'PinY', 'Angle']) {
		expect(attribute(after.get(name), 'V')).toBe(attribute(before.get(name), 'V'));
		expect(attribute(after.get(name), 'F')).toBe(attribute(before.get(name), 'F'));
	}
	const moved = await editVsdx(resized.bytes, [move]);
	expect(Number(attribute((await savedCells(moved.bytes)).get('Width'), 'V'))).toBe(4);
	const restored = await editVsdx(moved.bytes, [
		{ type: 'resize-shape', pageId: '0', shapeId: '1', width: 2, height: 0 },
	]);
	expect(Number(attribute((await savedCells(restored.bytes)).get('LocPinX'), 'V'))).toBe(1);
});

it('refuses locked/guarded width edits, nonzero line Height and zero 2D Height', async () => {
	const resize = { type: 'resize-shape' as const, pageId: '0', shapeId: '1', width: 4, height: 0 };
	for (const overrides of [{ LockWidth: [1] }, { Width: [2, 'GUARD(2)'] }] satisfies Record<
		string,
		[number, string?]
	>[])
		await expect(editVsdx(await source(line(3, 1.5, overrides)), [resize])).rejects.toThrow();
	await expect(editVsdx(await source(), [{ ...resize, height: 1 }])).rejects.toThrow('zero Height');
	const rectangleBytes = await source(shape('1', cell('Width', 2) + cell('Height', 1) + rectangle));
	await expect(editVsdx(rectangleBytes, [resize])).rejects.toThrow('positive Height');
});

it.each(['begin', 'end'] as const)(
	'edits and restores the %s endpoint with native formulas',
	async (which) => {
		const bytes = await source();
		const before = await savedCells(bytes);
		const saved = await editVsdx(bytes, [{ ...endpoint, endpoint: which }]);
		const after = await savedCells(saved.bytes);
		for (const name of ['Width', 'PinX', 'PinY', 'Angle', 'LocPinX', 'LocPinY'])
			expect(attribute(after.get(name), 'F')).toBe(attribute(before.get(name), 'F'));
		const prefix = which === 'begin' ? 'Begin' : 'End';
		const restored = await editVsdx(saved.bytes, [
			{
				...endpoint,
				endpoint: which,
				x: Number(attribute(before.get(`${prefix}X`), 'V')),
				y: Number(attribute(before.get(`${prefix}Y`), 'V')),
			},
		]);
		for (const [name, cell] of await savedCells(restored.bytes))
			expect(Number(attribute(cell, 'V')), name).toBeCloseTo(
				Number(attribute(before.get(name), 'V')),
				12,
			);
	},
);

it.each([
	{ LockEnd: [1] },
	{ LockWidth: [1] },
	{ LockMoveX: [1] },
	{ EndX: [3, 'GUARD(3)'] },
	{ Angle: [0, '0'] },
	{ Width: [2] },
	{ Angle: [1, 'ATAN2(EndY-BeginY,EndX-BeginX)'] },
	{ LocPinX: [2, 'Width*0.5'] },
] satisfies Record<string, [number, string?]>[])(
	'refuses unproven/protected endpoint changes (%s)',
	async (overrides) => {
		await expect(editVsdx(await source(line(3, 1.5, overrides)), [endpoint])).rejects.toThrow();
	},
);

it('refuses coincident endpoints and glued lines atomically', async () => {
	await expect(editVsdx(await source(), [{ ...endpoint, x: 1, y: 1.5 }])).rejects.toThrow(
		'Coincident',
	);
	await expect(
		editVsdx(
			await source(line(), '', '<Connects><Connect FromSheet="1" ToSheet="2"/></Connects>'),
			[endpoint],
		),
	).rejects.toThrow('Glued');
});

it('keeps endpoint no-ops byte-identical and bounds endpoint commands', async () => {
	const bytes = await source();
	const saved = await editVsdx(bytes, [{ ...endpoint, x: 3, y: 1.5 }]);
	expect(saved.bytes).toEqual(bytes);
	expect(saved.changedParts).toEqual([]);
	await expect(editVsdx(bytes, [{ ...endpoint, x: Infinity }])).rejects.toThrow('finite');
});

it('refuses endpoint length changes when absolute geometry cannot recalculate', async () => {
	const geometry = section(
		'Geometry',
		row(1, 'MoveTo', cell('X', 0) + cell('Y', 0)) + row(2, 'LineTo', cell('X', 2) + cell('Y', 0)),
	);
	await expect(editVsdx(await source(line(3, 1.5, {}, geometry)), [endpoint])).rejects.toThrow(
		'dimension-dependent',
	);
});
