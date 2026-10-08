import { computeTableMergeCrossings } from 'ooxml-ui/pptx';
import { describe, expect, it, vi } from 'vitest';

import { TableResizeOverlayComponent } from './table-resize-overlay.component';

describe('table resize release', () => {
	it.each(['col', 'row'] as const)('does not emit for a stationary %s boundary click', (type) => {
		// Exercise the DOM release handler without Angular render scheduling.
		const overlay = Object.create(
			TableResizeOverlayComponent.prototype,
		) as TableResizeOverlayComponent;
		const columns = vi.fn();
		const rows = vi.fn();
		const handle = document.createElement('div');
		Object.assign(overlay, {
			drag: {
				type,
				index: 0,
				startPos: 20,
				handle,
				initialWidths: [0.5, 0.5],
				initialRowHeight: 40,
			},
			onMove: vi.fn(),
			onUp: vi.fn(),
			resizeColumns: { emit: columns },
			resizeRow: { emit: rows },
		});
		overlay['handleUp'](new MouseEvent('pointerup', { clientX: 20, clientY: 20 }) as PointerEvent);
		expect(columns).not.toHaveBeenCalled();
		expect(rows).not.toHaveBeenCalled();
		expect(overlay['drag']).toBeNull();
	});
});

describe('table resize handles around merged cells', () => {
	it('draws no row handle over a vertically merged cell', () => {
		const overlay = Object.create(
			TableResizeOverlayComponent.prototype,
		) as TableResizeOverlayComponent;
		const rows = [
			{ cells: [{ text: 'A', rowSpan: 2 }, { text: 'B' }] },
			{ cells: [{ text: '', vMerge: true }, { text: 'C' }] },
		];
		Object.assign(overlay, {
			columnWidths: () => [0.25, 0.75],
			rowBounds: () => [40],
			tableHeight: () => 80,
			crossings: () => computeTableMergeCrossings(rows, 2),
		});
		expect(overlay.rowSegments(0)).toStrictEqual([{ leftPct: 25, widthPct: 75 }]);
		expect(overlay.colSegments(0)).toStrictEqual([{ top: 0, height: 80 }]);
	});
});
