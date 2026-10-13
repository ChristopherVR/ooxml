// @vitest-environment jsdom
import { expect, it, vi } from 'vitest';
import { renderTable } from './print-table';
it('compensates measured snapped border widths and removes its hidden probe', () => {
	const original = window.getComputedStyle.bind(window);
	const style = vi
		.spyOn(window, 'getComputedStyle')
		.mockImplementation((element, pseudo) =>
			(element as HTMLElement).dataset.printBorderProbe
				? ({ borderLeftWidth: '1px' } as CSSStyleDeclaration)
				: original(element, pseudo),
		);
	try {
		const edge = { widthPx: 0.625, style: 'solid' as const, color: '#123456' };
		const table = renderTable(
			{
				kind: 'table',
				blockId: 'table',
				yPx: 0,
				heightPx: 20,
				rows: [
					{
						yPx: 0,
						heightPx: 20,
						repeated: false,
						cells: [[]],
						geometry: [
							{
								xPx: 0,
								widthPx: 100,
								paddingLeftPx: 0.625,
								paddingRightPx: 0.625,
								contentHeightPx: 16,
								borderInsetsReserved: true,
								borders: { left: edge, top: edge, right: edge },
							},
						],
					},
				],
			},
			100,
			() => document.createElement('div'),
		);
		const content = table.querySelector<HTMLElement>('.dve-print-cell-content')!;
		expect(content.style.left).toBe('-0.375px');
		expect(content.style.right).toBe('-0.375px');
		expect(content.style.top).toBe('-1px');
		expect(document.querySelector('[data-print-border-probe]')).toBeNull();
	} finally {
		style.mockRestore();
	}
});
