import { describe, expect, it } from 'vitest';
import { createWorksheet } from '../workbook';
import { createGridMetrics } from './metrics';
import { type RemoteRange, layoutRemoteSelections } from './remote-selections';

const range = (r1: number, c1: number, r2 = r1, c2 = c1) => ({
	start: { row: r1, col: c1 },
	end: { row: r2, col: c2 },
});
const peer = (clientId: number, sheet: number, r: ReturnType<typeof range>): RemoteRange => ({
	clientId,
	userName: `User ${clientId}`,
	userColor: '#2563eb',
	sheet,
	range: r,
});

describe('layoutRemoteSelections', () => {
	const metrics = createGridMetrics(createWorksheet('Sheet1', 1));
	const options = { sheet: 0, maxRow: 200, maxCol: 50 };

	it('outlines peers on the shown sheet in plane pixels, ordered by client id', () => {
		const boxes = layoutRemoteSelections(
			metrics,
			[peer(9, 0, range(3, 1, 4, 2)), peer(2, 1, range(0, 0)), peer(4, 0, range(0, 0))],
			options,
		);
		expect(boxes.map((box) => box.clientId)).toEqual([4, 9]);
		const [first, second] = boxes;
		expect(second).toMatchObject({
			x: metrics.colLeft(1),
			y: metrics.rowTop(3),
			w: metrics.colLeft(3) - metrics.colLeft(1),
			h: metrics.rowTop(5) - metrics.rowTop(3),
			color: '#2563eb',
			label: 'User 9',
			tag: 'above',
		});
		expect(first?.tag).toBe('inside');
	});

	it('clips whole columns, skips off-screen and hidden ranges and shortens long names', () => {
		const sheet = createWorksheet('Sheet1', 1);
		sheet.columns.push({ min: 5, max: 5, width: 9, hidden: true });
		const hidden = createGridMetrics(sheet);
		const long = { ...peer(1, 0, range(0, 0, 1_048_575, 0)), userName: 'A'.repeat(40) };
		const boxes = layoutRemoteSelections(
			hidden,
			[long, peer(2, 0, range(300, 0)), peer(3, 0, range(0, 5))],
			{ ...options, maxLabelChars: 10 },
		);
		expect(boxes).toHaveLength(1);
		expect(boxes[0]).toMatchObject({ label: 'AAAAAAA...', name: 'A'.repeat(40) });
		expect(boxes[0]!.h).toBe(hidden.rowTop(201));
	});
});
