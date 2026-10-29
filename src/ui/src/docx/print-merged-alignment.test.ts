// @vitest-environment jsdom
import { expect, it } from 'vitest';
import { renderTable } from './print-table';

it('draws text aligned across a vertical merge below its first row, unclipped', () => {
	const paragraph = { kind: 'paragraph' as const, blockId: 'm', yPx: 0, heightPx: 20, lines: [] };
	const geometry = (extra = {}) => ({
		xPx: 0,
		widthPx: 100,
		paddingLeftPx: 0,
		paddingRightPx: 0,
		contentHeightPx: 20,
		...extra,
	});
	const table = renderTable(
		{
			kind: 'table',
			blockId: 't',
			yPx: 0,
			heightPx: 100,
			rows: [
				{
					yPx: 0,
					heightPx: 30,
					repeated: false,
					cells: [[paragraph]],
					geometry: [
						geometry({ verticalAlign: 'center' as const, mergeRows: 3, alignmentHeightPx: 100 }),
					],
				},
				{ yPx: 30, heightPx: 70, repeated: false, cells: [[]], geometry: [geometry()] },
			],
		},
		100,
		() => document.createElement('span'),
	);
	const [first, second] = Array.from(table.querySelectorAll<HTMLElement>('.dve-print-row'));
	const cell = first!.querySelector<HTMLElement>('.dve-print-cell')!;
	expect(cell.style.overflow).toBe('visible');
	expect(first!.style.zIndex).toBe('1');
	expect(cell.querySelector<HTMLElement>('.dve-print-cell-content')!.style.top).toBe('40px');
	expect(second!.querySelector<HTMLElement>('.dve-print-cell')!.style.overflow).toBe('');
});
