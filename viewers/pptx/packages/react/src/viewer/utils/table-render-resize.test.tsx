// @vitest-environment happy-dom
import React, { act } from 'react';
import { createRoot } from 'react-dom/client';
import { describe, expect, it, vi } from 'vitest';

import { TableResizeOverlay } from './table-render-resize';

describe('table resize release', () => {
	it.each(['col', 'row'])('does not commit a stationary %s boundary click', (axis) => {
		const host = document.createElement('div');
		document.body.appendChild(host);
		const root = createRoot(host);
		const columns = vi.fn();
		const rows = vi.fn();
		act(() =>
			root.render(
				<TableResizeOverlay
					columnWidths={[0.5, 0.5]}
					editable
					onResizeColumns={columns}
					onResizeRow={rows}
				>
					<table>
						<tbody>
							<tr>
								<td>A</td>
							</tr>
							<tr>
								<td>B</td>
							</tr>
						</tbody>
					</table>
				</TableResizeOverlay>,
			),
		);
		try {
			const handle = host.querySelector<HTMLElement>(`[class*="cursor-${axis}-resize"]`)!;
			act(() => {
				handle.dispatchEvent(
					new MouseEvent('mousedown', { bubbles: true, clientX: 20, clientY: 20 }),
				);
				document.dispatchEvent(new MouseEvent('mouseup', { clientX: 20, clientY: 20 }));
			});
			expect(columns).not.toHaveBeenCalled();
			expect(rows).not.toHaveBeenCalled();
			expect(document.body.style.cursor).toBe('');
		} finally {
			act(() => root.unmount());
			host.remove();
		}
	});
});

describe('table resize handles around merged cells', () => {
	it('leaves no row handle over a vertically merged cell', () => {
		const host = document.createElement('div');
		document.body.appendChild(host);
		const root = createRoot(host);
		act(() =>
			root.render(
				<TableResizeOverlay
					columnWidths={[0.25, 0.75]}
					rows={[
						{ cells: [{ text: 'A', rowSpan: 2 }, { text: 'B' }] },
						{ cells: [{ text: '', vMerge: true }, { text: 'C' }] },
					]}
					editable
				>
					<table>
						<tbody>
							<tr>
								<td rowSpan={2}>A</td>
								<td>B</td>
							</tr>
							<tr>
								<td>C</td>
							</tr>
						</tbody>
					</table>
				</TableResizeOverlay>,
			),
		);
		try {
			const segments = [...host.querySelectorAll<HTMLElement>('[class*="cursor-row-resize"]')];
			expect(segments.map((segment) => [segment.style.left, segment.style.width])).toStrictEqual([
				['25%', '75%'],
			]);
		} finally {
			act(() => root.unmount());
			host.remove();
		}
	});
});
