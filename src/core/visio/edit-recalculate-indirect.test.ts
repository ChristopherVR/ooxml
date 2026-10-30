import { describe, expect, it } from 'vitest';
import { parseXml } from '../xml/index';
import { VISIO_NS } from './sheet';
import { recalculateVisioCells } from './edit-recalculate';

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

describe('literal INDIRECT in the dependency graph', () => {
	it('treats INDIRECT("Sheet.1!Width") as a static dependency', () => {
		const root = page(
			shape('1', cell('Width', '8')) +
				shape('2', cell('PinX', '1', 'INDIRECT(&quot;Sheet.1!Width&quot;)*0.5')),
		);
		recalculateVisioCells(new Map([['0', root]]), changed);
		expect(valueOf(root, 'PinX')).toBe('4');
	});
	it('resolves a computed address while recalculating', () => {
		const root = page(
			shape('1', cell('Width', '8')) +
				shape('2', cell('PinX', '1', 'INDIRECT(&quot;Sheet.&quot;&amp;&quot;1!Width&quot;)*0.5')),
		);
		recalculateVisioCells(new Map([['0', root]]), changed);
		expect(valueOf(root, 'PinX')).toBe('4');
	});
	it('refuses a computed address that names no cell', () => {
		const root = page(
			shape('1', cell('Width', '8')) +
				shape('2', cell('PinX', '1', 'INDIRECT(&quot;Sheet.&quot;&amp;&quot;9!Width&quot;)')),
		);
		expect(() => recalculateVisioCells(new Map([['0', root]]), changed)).toThrow();
	});
});
