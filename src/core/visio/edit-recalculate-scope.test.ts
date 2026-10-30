import { describe, expect, it } from 'vitest';
import { parseXml } from '../xml/index';
import { VISIO_NS } from './sheet';
import { recalculateVisioCells, assertVisioShapeUnreferenced } from './edit-recalculate';
import { analyzeVisioFormula, parseVisioFormula } from './formula';

const page = (shapes: string) =>
	parseXml(`<PageContents xmlns="${VISIO_NS}"><Shapes>${shapes}</Shapes></PageContents>`)
		.documentElement;
const shape = (id: string, content: string) => `<Shape ID="${id}">${content}</Shape>`;
const cell = (name: string, value: string, formula = '') =>
	`<Cell N="${name}" V="${value}" ${formula ? `F="${formula}"` : ''}/>`;
const changed = [{ pageId: '0', shapeId: '1', cell: 'Width' }];
const valueOf = (root: Element, name: string) =>
	Array.from(root.getElementsByTagNameNS(VISIO_NS, 'Cell'))
		.find((node) => node.getAttribute('N') === name)!
		.getAttribute('V');

describe('formulas Visio writes in stencils and containers', () => {
	it('parses sheet-qualified calls, member reads of a returned sheet and omitted arguments', () => {
		expect(() =>
			parseVisioFormula('SUM(DEPENDSON(1,Sheet.2!SheetRef(),Sheet.1!SheetRef()))'),
		).not.toThrow();
		expect(() =>
			parseVisioFormula('IFERROR(CONTAINERSHEETREF(1,"Swimlane")!User.TEXT,"")'),
		).not.toThrow();
		expect(() => parseVisioFormula('SETATREF(User.FillForegndIn,,1)')).not.toThrow();
		expect(() => parseVisioFormula('LUM(ARG("C"))')).not.toThrow();
	});
	it('tells formulas that name what they read from those that may read anything', () => {
		const declared = analyzeVisioFormula('SUM(DEPENDSON(1,Sheet.2!SheetRef()))');
		expect(declared.dynamic).toBe(true);
		expect(declared.opaque).toBe(false);
		expect(analyzeVisioFormula('SETATREF(User.A,,1)+Width').opaque).toBe(false);
		expect(analyzeVisioFormula('EVALCELL("Width")').opaque).toBe(true);
		expect(analyzeVisioFormula('CONTAINERSHEETREF(1)!User.X').opaque).toBe(true);
		expect(analyzeVisioFormula('UNKNOWNFN(1)').opaque).toBe(true);
		expect(analyzeVisioFormula('ARG("C")').opaque).toBe(false);
	});
});

describe('scoped admission', () => {
	it('recalculates around a formula that names what it reads and is not affected', () => {
		const root = page(
			shape('1', cell('Width', '8') + cell('PinX', '1', 'Width*0.5')) +
				shape('2', cell('User', '0', 'SUM(DEPENDSON(1,Sheet.1!SheetRef()))')),
		);
		recalculateVisioCells(new Map([['0', root]]), changed);
		expect(valueOf(root, 'PinX')).toBe('4');
	});
	it('keeps the cache, naming the cell, when an edit reaches an unevaluable formula', () => {
		const root = page(
			shape('1', cell('Width', '8') + cell('PinX', '1', 'Width*0.5')) +
				shape('2', cell('LocPinX', '3', 'SETATREF(Sheet.1!Width)')),
		);
		const diagnostics: { code: string; message: string }[] = [];
		recalculateVisioCells(new Map([['0', root]]), changed, { diagnostics });
		expect(valueOf(root, 'PinX')).toBe('4');
		expect(valueOf(root, 'LocPinX')).toBe('3');
		expect(diagnostics).toEqual([
			expect.objectContaining({
				code: 'FORMULA_STALE',
				message: expect.stringContaining('LocPinX'),
			}),
		]);
	});
	it('assumes an opaque formula is reached by any edit on its page only', () => {
		const opaque = shape('2', cell('LocPinX', '0', 'EVALCELL(&quot;Width&quot;)'));
		const same = page(shape('1', cell('Width', '8')) + opaque);
		expect(() => recalculateVisioCells(new Map([['0', same]]), changed)).toThrow();
		const other = page(shape('1', cell('Width', '8')));
		const elsewhere = page(shape('1', cell('Width', '8')) + opaque);
		expect(() =>
			recalculateVisioCells(
				new Map([
					['0', other],
					['9', elsewhere],
				]),
				changed,
			),
		).not.toThrow();
	});
	it('refuses a deletion only when a retained opaque formula could read the shape', () => {
		const root = page(
			shape('1', cell('Width', '8')) +
				shape('2', cell('LocPinX', '0', 'EVALCELL(&quot;Width&quot;)')),
		);
		expect(() => assertVisioShapeUnreferenced(new Map([['0', root]]), '0', '1')).toThrow(
			/dynamic dependency/,
		);
	});
});
