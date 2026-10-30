import { describe, expect, it } from 'vitest';
import { buildXml, parseXml } from '../xml/index';
import { VISIO_NS } from './sheet';
import { recalculateVisioCells } from './edit-recalculate';

const page = (shapes: string) =>
	parseXml(`<PageContents xmlns="${VISIO_NS}"><Shapes>${shapes}</Shapes></PageContents>`)
		.documentElement;
const shape = (id: string, content: string, attrs = '') =>
	`<Shape ID="${id}" ${attrs}>${content}</Shape>`;
const cell = (name: string, value: string, formula = '') =>
	`<Cell N="${name}" V="${value}" ${formula ? `F="${formula}"` : ''}/>`;
const changed = [{ pageId: '0', shapeId: '1', cell: 'Width' }];
const cellOf = (root: Element, shapeId: string, name: string): Element => {
	const owner = Array.from(root.getElementsByTagNameNS(VISIO_NS, 'Shape')).find(
		(node) => node.getAttribute('ID') === shapeId,
	)!;
	return Array.from(owner.getElementsByTagNameNS(VISIO_NS, 'Cell')).find(
		(node) => node.getAttribute('N') === name,
	)!;
};
const run = (root: Element, options = {}) => {
	const diagnostics: { code: string; message: string }[] = [];
	recalculateVisioCells(new Map([['0', root]]), changed, { diagnostics, ...options });
	return diagnostics;
};

describe('dynamic references are resolved while recalculating', () => {
	it('follows EVALCELL given a cell, a text address and a computed address', () => {
		const root = page(
			shape('1', cell('Width', '8')) +
				shape(
					'2',
					cell('PinX', '0', 'EVALCELL(Sheet.1!Width)*0.5') +
						cell('PinY', '0', 'EVALCELL(&quot;Sheet.&quot;&amp;&quot;1!Width&quot;)*0.25'),
				),
		);
		run(root);
		expect(cellOf(root, '2', 'PinX').getAttribute('V')).toBe('4');
		expect(cellOf(root, '2', 'PinY').getAttribute('V')).toBe('2');
		expect(cellOf(root, '2', 'PinX').getAttribute('F')).toBe('EVALCELL(Sheet.1!Width)*0.5');
	});
	it('uses the EVALCELL fallback when the address names no value', () => {
		const root = page(
			shape('1', cell('Width', '8')) +
				shape('2', cell('PinX', '0', 'EVALCELL(&quot;1+&quot;,Sheet.1!Width/2)')),
		);
		run(root);
		expect(cellOf(root, '2', 'PinX').getAttribute('V')).toBe('4');
	});
	it('follows a chain of dynamic reads', () => {
		const root = page(
			shape('1', cell('Width', '8')) +
				shape(
					'2',
					cell('PinX', '0', 'INDIRECT(&quot;Sheet.&quot;&amp;1&amp;&quot;!Width&quot;)') +
						cell('PinY', '0', 'EVALCELL(PinX)*2'),
				),
		);
		run(root);
		expect(cellOf(root, '2', 'PinY').getAttribute('V')).toBe('16');
	});
	it('protects against a dynamic cycle', () => {
		const root = page(
			shape('1', cell('Width', '8')) +
				shape(
					'2',
					cell('PinX', '0', 'INDIRECT(&quot;Pi&quot;&amp;&quot;nY&quot;)') +
						cell('PinY', '0', 'PinX+1'),
				),
		);
		expect(() => run(root)).toThrow(/cycle/);
		const diagnostics = run(root, { cycles: 'keep' });
		expect(diagnostics.map((d) => d.code)).toContain('FORMULA_CYCLE');
	});
	it('refuses a dynamic address that names no cell', () => {
		const root = page(
			shape('1', cell('Width', '8')) +
				shape('2', cell('PinX', '0', 'INDIRECT(&quot;Sheet.&quot;&amp;&quot;7!Width&quot;)')),
		);
		expect(() => run(root)).toThrow();
	});
});

describe('formulas the evaluator cannot run keep their cache', () => {
	it('keeps an unknown function, reports it stale and leaves the formula untouched', () => {
		const root = page(
			shape('1', cell('Width', '8')) + shape('2', cell('PinX', '3', 'NOSUCHFN(Sheet.1!Width)')),
		);
		const diagnostics = run(root);
		expect(cellOf(root, '2', 'PinX').getAttribute('V')).toBe('3');
		expect(cellOf(root, '2', 'PinX').getAttribute('F')).toBe('NOSUCHFN(Sheet.1!Width)');
		expect(diagnostics.map((d) => d.code)).toEqual(['FORMULA_STALE']);
	});
	it('does not report an unknown function the edit cannot reach', () => {
		const root = page(
			shape('1', cell('Width', '8') + cell('PinX', '1', 'Width*0.5')) +
				shape('2', cell('PinX', '3', 'NOSUCHFN(Height)')),
		);
		expect(run(root)).toEqual([]);
	});
	it('follows DEPENDSON as an explicit dependency and keeps its cache', () => {
		const root = page(
			shape('1', cell('Width', '8')) +
				shape('2', cell('User', '1', 'DEPENDSON(1,Sheet.1!Width)') + cell('PinX', '1', 'User*2')),
		);
		const diagnostics = run(root);
		expect(diagnostics.map((d) => d.code)).toEqual(['FORMULA_STALE']);
		expect(cellOf(root, '2', 'PinX').getAttribute('V')).toBe('2');
	});
	it('a SheetRef() formula survives unchanged when an unrelated edit recalculates', () => {
		const formula = 'SUM(DEPENDSON(1,Sheet.1!SheetRef()))';
		const root = page(
			shape('1', cell('Width', '8') + cell('PinX', '1', 'Width*0.5')) +
				shape('2', cell('User', '0', formula)),
		);
		run(root);
		expect(cellOf(root, '2', 'User').getAttribute('F')).toBe(formula);
		expect(cellOf(root, '1', 'PinX').getAttribute('V')).toBe('4');
	});
	it('still refuses drawn geometry data it cannot recompute', () => {
		const root = page(
			shape('1', cell('Width', '8')) + shape('2', cell('X1', '3', 'POLYLINE(0,0,Sheet.1!Width,1)')),
		);
		const before = buildXml(root.ownerDocument!);
		expect(() => run(root)).toThrow();
		expect(buildXml(root.ownerDocument!)).toBe(before);
	});
});

describe('reads of grouped and instanced shapes', () => {
	it('reads a cell of a shape nested in a group as an input the edit does not reach', () => {
		const root = page(
			shape('1', cell('Width', '8')) +
				shape(
					'2',
					cell('Width', '4') + `<Shapes>${shape('3', cell('Height', '6'))}</Shapes>`,
					'Type="Group"',
				) +
				shape('4', cell('PinX', '0', 'Sheet.3!Height+Sheet.2!Width+Sheet.1!Width')),
		);
		run(root);
		expect(cellOf(root, '4', 'PinX').getAttribute('V')).toBe('18');
	});
	it('reads a locally set cell of an instanced shape', () => {
		const root = page(
			shape('1', cell('Width', '8')) +
				shape('2', cell('Width', '5'), 'Master="9"') +
				shape('3', cell('PinX', '0', 'Sheet.2!Width+Sheet.1!Width')),
		);
		run(root);
		expect(cellOf(root, '3', 'PinX').getAttribute('V')).toBe('13');
	});
	it('still refuses when the edit itself reaches a grouped cell', () => {
		const root = page(
			shape('1', cell('Width', '8')) +
				shape('2', cell('Width', '4', 'Sheet.1!Width'), 'Type="Group"'),
		);
		expect(() => run(root)).toThrow();
	});
	it('still refuses a read of an inherited formula it cannot see', () => {
		const root = page(
			shape('1', cell('Width', '8')) +
				shape('2', '<Cell N="Width" V="5" F="Inh"/>', 'Master="9"') +
				shape('3', cell('PinX', '0', 'Sheet.2!Width+Sheet.1!Width')),
		);
		expect(() => run(root)).toThrow();
	});
});
