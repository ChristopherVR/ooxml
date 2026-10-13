import type { RowLayout } from './table-layout';
import { cellContentBottom } from './table-content';
/** Native two/three-column probes finish a consumed leaf story's last line beside a later
 * enclosing marker. Longer stories stay at the reference band, rather than filling note slack. */
export function completedNoteStoryMinimum(
	row: RowLayout,
	stories: ReadonlyMap<string, string>,
	pendingBlocks: ReadonlySet<string>,
	available: number,
): number {
	let minimum = 0;
	for (const [index, cell] of row.cells.entries()) {
		const groups = new Map<string, typeof cell>();
		for (const box of cell) {
			const group = stories.get(box.blockId);
			if (group !== undefined) {
				const boxes = groups.get(group) ?? [];
				boxes.push(box);
				groups.set(group, boxes);
			}
		}
		for (const boxes of groups.values()) {
			if (
				boxes.reduce((count, box) => count + box.lines.length, 0) !== 1 ||
				boxes.some((box) => pendingBlocks.has(box.blockId))
			)
				continue;
			const bottom = cellContentBottom(boxes, row.cellPadding?.[index]?.bottom ?? 0);
			if (bottom <= available + 1e-6) minimum = Math.max(minimum, bottom);
		}
	}
	return minimum;
}
