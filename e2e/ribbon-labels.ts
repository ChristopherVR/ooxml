/** Checks the painted text, including wrapped captions and generated group labels. */
export function clippedRibbonLabels(panel: Element): string[] {
	const clipped: string[] = [];
	if (panel.scrollWidth > panel.clientWidth + 1) clipped.push('panel scrolls horizontally');
	const context = document.createElement('canvas').getContext('2d');
	for (const group of panel.querySelectorAll<HTMLElement>(':scope > .ribbon-group')) {
		if (!group.getClientRects().length) continue;
		if (!group.hasAttribute('data-collapsed') && context && group.dataset.caption) {
			const style = getComputedStyle(group, '::after');
			context.font = `${style.fontSize} ${style.fontFamily}`;
			if (context.measureText(group.dataset.caption).width > group.clientWidth + 1)
				clipped.push(group.dataset.caption);
		}
	}
	const range = document.createRange();
	for (const caption of panel.querySelectorAll<HTMLElement>(
		'.ribbon-large > span, .ribbon-menu > span, .ribbon-caption',
	)) {
		if (!caption.getClientRects().length || !caption.textContent?.trim()) continue;
		range.selectNodeContents(caption);
		const owner = caption.closest('button, .ribbon-menu') ?? caption;
		const outer = owner.getBoundingClientRect();
		const inner = caption.getBoundingClientRect();
		const outside =
			[...range.getClientRects()].some((box) => {
				// Font ascent and descent can extend past a tightly spaced line box. Compare the
				// text to the control vertically, and use the caption's own overflow metrics below.
				return (
					box.left < Math.max(outer.left, inner.left) - 1 ||
					box.right > Math.min(outer.right, inner.right) + 1 ||
					box.top < outer.top - 1 ||
					box.bottom > outer.bottom + 1
				);
			}) || caption.scrollHeight > caption.clientHeight + 1;
		if (outside) clipped.push(caption.textContent);
	}
	return clipped;
}
