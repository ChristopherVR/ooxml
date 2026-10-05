import { describe, expect, it } from 'vitest';
import { parseXml, buildXml } from '../xml/index.js';
import { VISIO_NS } from './sheet.js';
import {
	recalculateVisioCells,
	assertVisioShapeUnreferenced,
	visioCellDependsOn,
} from './edit-recalculate.js';
const page = (shapes: string, extra = '') =>
	parseXml(`<PageContents xmlns="${VISIO_NS}"><Shapes>${shapes}</Shapes>${extra}</PageContents>`)
		.documentElement;
const shape = (id: string, content: string, attrs = '') =>
	`<Shape ID="${id}" ${attrs}>${content}</Shape>`;
const cell = (name: string, value: string, formula = '', attrs = '') =>
	`<Cell N="${name}" V="${value}" ${formula ? `F="${formula}"` : ''} ${attrs}/>`;
const widthChange = [{ pageId: '0', shapeId: '1', cell: 'Width' }];
const getCell = (root: Element, name: string) =>
	Array.from(root.getElementsByTagNameNS(VISIO_NS, 'Cell')).find(
		(node) => node.getAttribute('N') === name,
	)!;
describe('dependency-aware ShapeSheet cache writeback', () => {
	it('preserves independent themed caches and rejects affected implicit theme inputs atomically', () => {
		const root = page(
			shape(
				'1',
				cell('Width', '8') +
					cell('Height', '4') +
					cell('ColorSchemeIndex', '1') +
					cell('FillForegnd', '#ffffff', 'THEMEVAL()'),
			),
		);
		recalculateVisioCells(new Map([['0', root]]), widthChange);
		expect(getCell(root, 'FillForegnd').getAttribute('V')).toBe('#ffffff');
		getCell(root, 'ColorSchemeIndex').setAttribute('F', 'Width/Height');
		const before = buildXml(root.ownerDocument!);
		expect(() => recalculateVisioCells(new Map([['0', root]]), widthChange)).toThrow(
			/unsupported functions/,
		);
		expect(buildXml(root.ownerDocument!)).toBe(before);
	});
	it('rejects affected connected, one-dimensional and foreign shapes', () => {
		for (const [contents, attrs, extra] of [
			['', '', '<Connects><Connect FromSheet="2" ToSheet="3"/></Connects>'],
			[cell('OneD', '1'), '', ''],
			[cell('BeginX', '0'), '', ''],
			['<ForeignData/>', '', ''],
			['', 'Type="Foreign"', ''],
		] as const) {
			const root = page(
				shape('1', cell('Width', '8')) +
					shape('2', cell('PinX', '1', 'Sheet.1!Width') + contents, attrs),
				extra,
			);
			expect(() => recalculateVisioCells(new Map([['0', root]]), widthChange)).toThrow(
				/unsupported|missing, inherited or grouped/,
			);
		}
	});
	it('recalculates formulas referencing implicit text dimensions and text pins', () => {
		const root = page(
			shape(
				'1',
				cell('Width', '8') +
					cell('Height', '4') +
					cell('TxtPinX', '1', 'TxtWidth*0.5') +
					cell('TxtLocPinY', '1', 'TxtHeight*0.5'),
			),
		);
		recalculateVisioCells(new Map([['0', root]]), [
			...widthChange,
			{ pageId: '0', shapeId: '1', cell: 'Height' },
		]);
		expect(getCell(root, 'TxtPinX').getAttribute('V')).toBe('4');
		expect(getCell(root, 'TxtLocPinY').getAttribute('V')).toBe('2');
	});
	it('rejects formula-bearing markup with unknown cell scope', () => {
		const root = page(
			shape('1', cell('Width', '8')),
			'<Extension><Cell N="Odd" F="Sheet.1!Width" V="2"/></Extension>',
		);
		expect(() => recalculateVisioCells(new Map([['0', root]]), widthChange)).toThrow(
			/outside admitted/,
		);
	});
	it('recalculates guarded geometry through static cross-shape dependencies and synthetic LocPin defaults', () => {
		const root = page(
			shape('1', cell('Width', '8')) +
				shape(
					'2',
					cell('PinX', '1', 'Sheet.1!LocPinX', 'U="DL"') +
						`<Section N="Geometry" IX="0"><Row IX="1" T="LineTo">${cell('X', '1', 'GUARD(PinX*0.5)', 'U="DL"')}${cell('Y', '1')}</Row><Row IX="2" T="LineTo">${cell('X', '1', 'Geometry1.X1*2', 'U="DL"')}</Row></Section>`,
				),
		);
		expect(recalculateVisioCells(new Map([['0', root]]), widthChange)).toEqual(['0']);
		expect(getCell(root, 'PinX').getAttribute('V')).toBe('4');
		const xs = Array.from(root.getElementsByTagNameNS(VISIO_NS, 'Cell')).filter(
			(node) => node.getAttribute('N') === 'X',
		);
		expect(xs.map((node) => node.getAttribute('V'))).toEqual(['2', '4']);
		expect(xs[0]!.getAttribute('F')).toBe('GUARD(PinX*0.5)');
		expect(xs[0]!.getAttribute('U')).toBe('DL');
	});
	it('preserves unsupported independent functions while rejecting affected functions atomically', () => {
		const root = page(
			shape(
				'1',
				cell('Width', '8') +
					cell('Height', '2') +
					cell('PinX', '1', 'Width*0.5') +
					cell('PinY', '2', 'ROUND(Height,0)'),
			),
		);
		recalculateVisioCells(new Map([['0', root]]), widthChange);
		expect(getCell(root, 'PinY').getAttribute('V')).toBe('2');
		getCell(root, 'PinY').setAttribute('F', 'ROUND(Width,0)');
		getCell(root, 'PinX').setAttribute('V', '1');
		const before = buildXml(root.ownerDocument!);
		expect(() => recalculateVisioCells(new Map([['0', root]]), widthChange)).toThrow(
			/unsupported functions/,
		);
		expect(buildXml(root.ownerDocument!)).toBe(before);
	});
	it.each(['DEPENDSON(Width)', 'SETATREF(Width)', 'unparseable@Width'])(
		'rejects unknown dependency behavior: %s',
		(formula) => {
			const root = page(shape('1', cell('Width', '8') + cell('User', '1', formula)));
			expect(() => recalculateVisioCells(new Map([['0', root]]), widthChange)).toThrow(
				/dependency|dependencies/,
			);
		},
	);
	it('rejects affected cycles and bad units without partial cache writes', () => {
		const root = page(
			shape('1', cell('Width', '8') + cell('PinX', '1', 'Width+PinY') + cell('PinY', '1', 'PinX')),
		);
		const before = buildXml(root.ownerDocument!);
		expect(() => recalculateVisioCells(new Map([['0', root]]), widthChange)).toThrow(/cycle/);
		expect(buildXml(root.ownerDocument!)).toBe(before);
		getCell(root, 'PinY').setAttribute('F', 'Width*1 rad');
		expect(() => recalculateVisioCells(new Map([['0', root]]), widthChange)).toThrow();
	});
	it('rejects affected error caches, master cells, and grouped cells', () => {
		for (const attrs of ['Master="2"', 'Type="Group"']) {
			const root = page(
				shape('1', cell('Width', '8')) +
					shape(
						'2',
						cell('PinX', '1', 'Sheet.1!Width') +
							(attrs.includes('Group') ? '<Shapes><Shape ID="3"/></Shapes>' : ''),
						attrs,
					),
			);
			expect(() => recalculateVisioCells(new Map([['0', root]]), widthChange)).toThrow(
				/inherited or grouped/,
			);
		}
		const root = page(shape('1', cell('Width', '8') + cell('PinX', '1', 'Width', 'E="#VALUE!"')));
		expect(() => recalculateVisioCells(new Map([['0', root]]), widthChange)).toThrow(/error cache/);
	});
	it('enforces cell, closure and evaluation limits', () => {
		const root = page(
			shape('1', cell('Width', '8') + cell('PinX', '1', 'Width') + cell('PinY', '1', 'PinX')),
		);
		for (const options of [{ maxCells: 1 }, { maxAffectedCells: 1 }, { maxSteps: 1 }])
			expect(() => recalculateVisioCells(new Map([['0', root]]), widthChange, options)).toThrow(
				/limit/,
			);
	});
	it('proves deletion unreferenced across formula cells and Connect records', () => {
		const root = page(
			shape('1', cell('Width', '8')) + shape('2', cell('PinX', '1', 'Sheet.1!Width')),
		);
		expect(() => assertVisioShapeUnreferenced(new Map([['0', root]]), '0', '1')).toThrow(
			/referenced/,
		);
		expect(() => assertVisioShapeUnreferenced(new Map([['0', root]]), '0', '2')).not.toThrow();
		const connected = page(
			shape('1', cell('Width', '8')),
			'<Connects><Connect FromSheet="1" ToSheet="2"/></Connects>',
		);
		expect(() => assertVisioShapeUnreferenced(new Map([['0', connected]]), '0', '1')).toThrow(
			/Connect/,
		);
	});
	it('keeps pages isolated even when shape IDs collide', () => {
		const first = page(shape('1', cell('Width', '8') + cell('PinX', '1', 'Width*0.5')));
		const second = page(shape('1', cell('Width', '10') + cell('PinX', '1', 'Width*0.5')));
		expect(
			recalculateVisioCells(
				new Map([
					['0', first],
					['1', second],
				]),
				widthChange,
			),
		).toEqual(['0']);
		expect(getCell(second, 'PinX').getAttribute('V')).toBe('1');
	});
	it('recalculates named User rows and dimensionless relative geometry', () => {
		const root = page(
			shape(
				'1',
				cell('Width', '8') +
					`<Section N="User"><Row N="half">${cell('Value', '1', 'Width*0.5', 'U="DL"')}</Row></Section><Section N="Geometry" IX="0"><Row IX="1" T="RelLineTo">${cell('X', '1', 'User.half/Width')}${cell('Y', '0')}</Row></Section>`,
			),
		);
		recalculateVisioCells(new Map([['0', root]]), widthChange);
		expect(getCell(root, 'Value').getAttribute('V')).toBe('4');
		expect(getCell(root, 'X').getAttribute('V')).toBe('0.5');
		expect(
			visioCellDependsOn(
				new Map([['0', root]]),
				{ pageId: '0', shapeId: '1', cell: 'Geometry1.X1' },
				widthChange,
			),
		).toBe(true);
		expect(
			visioCellDependsOn(
				new Map([['0', root]]),
				{ pageId: '0', shapeId: '1', cell: 'Geometry1.Y1' },
				widthChange,
			),
		).toBe(false);
	});
	it('bounds dependency depth and charges aggregate AST evaluation steps', () => {
		const root = page(
			shape(
				'1',
				cell('Width', '8') + cell('PinX', '1', 'Width*0.5') + cell('PinY', '1', 'PinX*0.5'),
			),
		);
		expect(() =>
			recalculateVisioCells(new Map([['0', root]]), [{ pageId: '0', shapeId: '1', cell: 'PinY' }], {
				maxDepth: 1,
			}),
		).toThrow(/depth limit/);
		expect(() =>
			recalculateVisioCells(new Map([['0', root]]), widthChange, { maxSteps: 5 }),
		).toThrow(/limit/);
	});
});
