import { parseInkMlContent } from '../../../drawingml/inkml/inkml-content-part';
import { readInkPart } from '../../../drawingml/inkml/part';
import { inkObjectFromXml } from '../../../drawingml/inkml/xml-tree';
import { inkFromClipboard } from '../../../docx/ui/ink-clipboard';
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
	{
		name: 'pasted Word ink is rebuilt from what ink may hold',
		limits: { ms: 1500, heapMb: 96 },
		run() {
			const dataset = {
				ink: '{"strokes":[{"path":"M 0 0 L 1 1","color":"#000000","width":1e999}],"width":1,"height":1}',
			};
			const huge = inkFromClipboard({ dataset } as unknown as HTMLElement);
			const placement = inkFromClipboard({
				dataset: {
					ink: '{"strokes":[{"path":"M 0 0 L 1 1","color":"#000000","width":1}],"width":1,"height":1}',
					anchored: '1',
					placement: '{"wrap":"x","offsetYPx":1e999}',
					widthPx: '1e999',
				},
			} as unknown as HTMLElement);
			return `${'ink' in huge} ${placement.placement} ${placement.widthPx}`;
		},
		expected: 'false {"wrap":"none"} 1',
	},
];
