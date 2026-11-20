import type { LayoutBlockBox, TextBoxLayout } from 'ooxml-core/docx/layout';

/** Draws one laid-out block at a column width, as a Print Layout page draws its body blocks. */
export type PrintBlock = (block: LayoutBlockBox, widthPx: number) => HTMLElement;

/**
 * A text box's lines, placed as the page engine laid them out (`TextBoxLayout`): clipped to the
 * shape's text rectangle as Word clips overflowing text, turned a quarter for vertical text, and in
 * `color` where the runs set none.
 */
export function textBoxLayer(
	layout: TextBoxLayout,
	color: string | undefined,
	renderBlock: PrintBlock,
): HTMLElement {
	const frame = document.createElement('div');
	frame.className = 'dve-print-text-box';
	Object.assign(frame.style, {
		left: `${layout.frame.xPx}px`,
		top: `${layout.frame.yPx}px`,
		width: `${layout.frame.widthPx}px`,
		height: `${layout.frame.heightPx}px`,
	});
	if (color && /^#[0-9a-f]{6}$/i.test(color)) frame.style.color = color;
	const length = layout.rotation ? layout.frame.heightPx : layout.frame.widthPx;
	const depth = layout.rotation ? layout.frame.widthPx : layout.frame.heightPx;
	const text = document.createElement('div');
	text.className = 'dve-print-text-box-text';
	text.style.width = `${length}px`;
	text.style.height = `${depth}px`;
	// Turned about its top-left corner: `vert` puts the first line at the right, `vert270` at the left.
	if (layout.rotation === 90) text.style.transform = `translate(${depth}px, 0) rotate(90deg)`;
	if (layout.rotation === 270) text.style.transform = `translate(0, ${length}px) rotate(-90deg)`;
	for (const block of layout.blocks) text.append(renderBlock(block, length));
	// The box's lines are not the story's: nothing may resolve them as editable blocks.
	for (const line of text.querySelectorAll<HTMLElement>('[data-block-id]'))
		delete line.dataset.blockId;
	frame.append(text);
	return frame;
}
