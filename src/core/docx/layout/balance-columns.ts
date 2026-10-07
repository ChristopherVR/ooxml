import type { LayoutPageBox } from './result.js';
import type { LayoutBlock } from './input.js';

/** Only layouts whose trial flow can preserve every indivisible unit are balanced. */
export function canBalanceBlock(block: LayoutBlock): boolean {
	if (block.kind === 'paragraph')
		return !block.floats?.length && !block.runs.some((run) => run.breakAfter);
	return (
		block.rows.length > 0 &&
		block.rows.every((row) =>
			row.cells.every((cell) =>
				cell.paragraphs.every(
					(paragraph) => !paragraph.footnotes?.length && canBalanceBlock(paragraph),
				),
			),
		)
	);
}

/** Reflow only the final page at a lower column capacity, retaining earlier pages' capacity. */
export function balanceColumns<T extends { balanceValid: boolean }>(
	basePages: LayoutPageBox[],
	pageIndex: number,
	fullHeightPx: number,
	run: (pages: LayoutPageBox[], heightPx: number) => T,
): { pages: LayoutPageBox[]; flow: T } {
	const attempt = (heightPx: number) => {
		const pages = structuredClone(basePages);
		const flow = run(pages, heightPx);
		return { pages, flow };
	};
	let low = 0;
	let high = fullHeightPx;
	let best = attempt(high);
	// Word capacities are in twips. Resolve below one twip without requiring a font-specific
	// line height or splitting a keep-together group at the trial page height.
	for (let step = 0; step < 16 && high - low > 1 / 15; step++) {
		const middle = (low + high) / 2;
		const candidate = attempt(middle);
		if (candidate.pages.length <= pageIndex + 1 && candidate.flow.balanceValid) {
			high = middle;
			best = candidate;
		} else low = middle;
	}
	return best;
}
