// @vitest-environment jsdom
import { expect, it } from 'vitest';
import { splitRowAtHeight } from 'ooxml-core/docx/layout';
import { renderTable } from './print-table';
it('paints a continuation cell tall enough for its longest parallel story', () => {
	const lines = (count: number) =>
		Array.from({ length: count }, (_, i) => ({
			yPx: i * 16,
			heightPx: 16,
			firstToken: i,
			lastToken: i + 1,
			sourceStart: i,
			sourceEnd: i + 1,
			fragments: [],
		}));
	const long = {
		kind: 'paragraph' as const,
		blockId: 'long',
		yPx: 4,
		heightPx: 64,
		tableFrameInsets: { top: 4, bottom: 4 },
		lines: lines(4),
	};
	const short = { ...long, blockId: 'short', heightPx: 48, lines: lines(3) };
	for (const paragraphs of [
		[long, short],
		[short, long],
	]) {
		const { after } = splitRowAtHeight(
			{
				heightPx: 72,
				cells: [paragraphs],
				geometry: [
					{ xPx: 0, widthPx: 160, paddingLeftPx: 0, paddingRightPx: 0, contentHeightPx: 72 },
				],
				cellPadding: [{ top: 2, bottom: 3 }],
			},
			27,
		);
		const table = renderTable(
			{
				kind: 'table',
				blockId: 'table',
				yPx: 0,
				heightPx: after!.heightPx,
				rows: [
					{
						yPx: 0,
						heightPx: after!.heightPx,
						repeated: false,
						cells: after!.cells,
						geometry: after!.geometry,
					},
				],
			},
			160,
			(paragraph) => {
				const element = document.createElement('div');
				element.dataset.story = paragraph.blockId;
				element.style.top = paragraph.yPx + 'px';
				element.style.height = paragraph.heightPx + 'px';
				return element;
			},
		);
		const cell = table.querySelector<HTMLElement>('.dve-print-cell')!;
		const ink = table.querySelector<HTMLElement>('[data-story="long"]')!;
		expect(parseFloat(cell.style.height)).toBe(61);
		expect(parseFloat(ink.style.top) + parseFloat(ink.style.height) + 7).toBe(
			parseFloat(cell.style.height),
		);
	}
});
