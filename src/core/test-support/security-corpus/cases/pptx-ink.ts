import { parseInkMlContent } from '../../../drawingml/inkml/inkml-content-part';
import { readInkPart } from '../../../drawingml/inkml/part';
import { inkObjectFromXml } from '../../../drawingml/inkml/xml-tree';
import type { SecurityCase } from '../case';

export const cases: readonly SecurityCase[] = [
	{
		name: 'ink traces beyond the per-stroke sample budget are omitted',
		limits: { ms: 1500, heapMb: 96 },
		run() {
			const trace = '0 0,' + '1 1,'.repeat(4096) + '2 2';
			return parseInkMlContent({ ink: { trace } }).strokes.length;
		},
		expected: 0,
	},
	{
		name: 'ink element names such as __proto__ stay ordinary entries of the tree',
		limits: { ms: 1500, heapMb: 96 },
		run() {
			const xml =
				'<inkml:ink xmlns:inkml="http://www.w3.org/2003/InkML"><__proto__ polluted="yes"/>' +
				'<constructor prototype="x"/><inkml:trace>0 0,10 10</inkml:trace></inkml:ink>';
			const strokes = readInkPart(xml).strokes.length;
			const root = inkObjectFromXml(xml)['inkml:ink'] as object;
			// Kept as ordinary entries, not as the tree's prototype.
			const own = Object.hasOwn(root, '__proto__') && Object.getPrototypeOf(root) === null;
			return `${strokes} ${own}`;
		},
		expected: '1 true',
	},
];
