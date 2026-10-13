import type { LayoutParagraphBox } from './result';
/** Parallel cell stories need the greatest bottom, independent of source ordering. */
export function cellContentBottom(boxes: readonly LayoutParagraphBox[], paddingBottom = 0): number {
	return Math.max(
		0,
		...boxes.map(
			(box) => box.yPx + box.heightPx + (box.tableFrameInsets?.bottom ?? 0) + paddingBottom,
		),
	);
}
