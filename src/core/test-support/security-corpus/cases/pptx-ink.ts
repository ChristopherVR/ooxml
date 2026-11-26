import { parseInkMlContent } from '../../../drawingml/inkml/inkml-content-part';
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
];
