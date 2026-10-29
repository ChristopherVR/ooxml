// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
// The `mc:Fallback` of a shape Word wrote for old readers (VML) is kept when the shape is edited.
// Only the width and height of its VML shapes follow a resize; every other edit (preset, fill,
// text, position) is not mirrored, so the fallback can be stale for Word 2007.
import type { InlineImage } from './model';
import { isElement, type XmlElement } from './xml';

const PT_PER_PX = 0.75;

function setStyleSize(style: string, name: 'width' | 'height', pt: number): string {
	const value = `${Math.round(pt * 100) / 100}pt`;
	const pattern = new RegExp(`(^|;)\\s*${name}\\s*:\\s*[^;]*`, 'i');
	return pattern.test(style)
		? style.replace(pattern, (_all, lead: string) => `${lead}${name}:${value}`)
		: style;
}

/** Updates the width and height in the VML styles of the fallback when the shape was resized. */
export function patchFallbackSize(
	content: XmlElement,
	image: InlineImage,
	base: InlineImage | undefined,
): void {
	if (!base || (image.widthPx === base.widthPx && image.heightPx === base.heightPx)) return;
	const fallback = Array.from(content.childNodes)
		.filter(isElement)
		.find((node) => node.localName === 'Fallback');
	if (!fallback) return;
	const stack: XmlElement[] = [fallback];
	while (stack.length) {
		const node = stack.pop()!;
		const style = node.getAttribute('style');
		if (style && /(^|;)\s*width\s*:/i.test(style)) {
			let next = style;
			if (image.widthPx !== base.widthPx)
				next = setStyleSize(next, 'width', image.widthPx * PT_PER_PX);
			if (image.heightPx !== base.heightPx)
				next = setStyleSize(next, 'height', image.heightPx * PT_PER_PX);
			node.setAttribute('style', next);
			continue;
		}
		for (const child of Array.from(node.childNodes)) if (isElement(child)) stack.push(child);
	}
}
