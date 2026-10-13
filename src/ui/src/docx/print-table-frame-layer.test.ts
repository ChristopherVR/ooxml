// @vitest-environment jsdom
import { expect, it } from 'vitest';
import { renderTable } from './print-table';
it('lets external frame edges paint horizontally while clipping exact-row height beneath ink', () => {
	const paragraph = { kind: 'paragraph' as const, blockId: 'ink', yPx: 0, heightPx: 16, lines: [] };
	const nested = {
		kind: 'table' as const,
		blockId: 'nested',
		xPx: 0,
		widthPx: 102,
		yPx: 0,
		heightPx: 40,
		rows: [{ yPx: 0, heightPx: 40, repeated: false, cells: [[]] }],
	};
	let painted = 0;
	const table = renderTable(
		{
			kind: 'table',
			blockId: 'outer',
			yPx: 0,
			heightPx: 20,
			rows: [
				{
					yPx: 0,
					heightPx: 20,
					repeated: false,
					cells: [[paragraph]],
					geometry: [
						{ xPx: 0, widthPx: 100, paddingLeftPx: 0, paddingRightPx: 0, contentHeightPx: 16 },
					],
					cellTables: [[nested]],
				},
			],
		},
		100,
		() => {
			painted++;
			return document.createElement('span');
		},
	);
	const row = table.querySelector<HTMLElement>('.dve-print-row')!,
		cell = row.querySelector<HTMLElement>(':scope > .dve-print-cell')!,
		layer = row.querySelector<HTMLElement>(':scope > .dve-print-cell-frame-layer')!;
	expect(row.firstElementChild).toBe(cell);
	expect(layer.closest('.dve-print-cell')).toBeNull();
	expect(row.style.isolation).toBe('isolate');
	expect(layer.style.height).toBe('20px');
	expect(layer.style.overflowY).toBe('clip');
	expect(layer.style.overflowX).toBe('visible');
	expect(layer.style.pointerEvents).toBe('none');
	expect(layer.style.zIndex).toBe('1');
	expect(cell.querySelector<HTMLElement>('.dve-print-cell-content')!.style.zIndex).toBe('2');
	expect(layer.textContent).toBe('');
	expect(painted).toBe(1);
});

it.each([false, true])(
	'preserves the ordinary or reserved enclosing border origin: %s',
	(reserved) => {
		const edge = { widthPx: 4, style: 'solid' as const, color: '#000000' };
		const nested = {
			kind: 'table' as const,
			blockId: 'nested',
			xPx: 0,
			widthPx: 82,
			yPx: 0,
			heightPx: 16,
			rows: [{ yPx: 0, heightPx: 16, repeated: false, cells: [[]] }],
		};
		const table = renderTable(
			{
				kind: 'table',
				blockId: 'outer',
				yPx: 0,
				heightPx: 24,
				rows: [
					{
						yPx: 0,
						heightPx: 24,
						repeated: false,
						cells: [[]],
						geometry: [
							{
								xPx: 10,
								widthPx: 100,
								paddingLeftPx: 6,
								paddingRightPx: 0,
								contentHeightPx: 24,
								borders: { left: edge, top: edge },
								borderInsetsReserved: reserved,
							},
						],
						cellTables: [[nested]],
					},
				],
			},
			100,
			() => document.createElement('div'),
		);
		const layer = table.querySelector<HTMLElement>('.dve-print-cell-frame-layer')!,
			origin = layer.querySelector<HTMLElement>('.dve-print-cell-frame-origin')!;
		expect(layer.style.left).toBe('10px');
		expect(layer.style.top).toBe('0px');
		expect(origin.style.left).toBe('6px');
		expect(origin.style.top).toBe('0px');
		expect(layer.style.borderLeftWidth).toBe(reserved ? '' : '4px');
		expect(layer.style.borderTopWidth).toBe(reserved ? '' : '4px');
		if (!reserved) expect(layer.style.borderLeftColor).toBe('transparent');
	},
);
