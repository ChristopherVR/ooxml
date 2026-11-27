/**
 * A large ribbon command's caption. Word breaks these only at spaces ("Get" / "Add-ins", never
 * "Get Add-" / "ins"), but a browser also breaks after a hyphen and no CSS property turns that
 * off, so each word sits in its own unbreakable span. `textContent` stays the caption, and the
 * English caption is kept so a locale change rewrites the words.
 */
export function writeLargeCaption(caption: HTMLElement, text: string, key = text): void {
	caption.classList.add('ribbon-caption');
	caption.dataset.captionKey = key;
	const words = text.split(/(\s+)/u);
	caption.replaceChildren(
		...words.map((word) => {
			if (!word || /^\s+$/u.test(word)) return caption.ownerDocument.createTextNode(word);
			const span = caption.ownerDocument.createElement('span');
			span.className = 'ribbon-caption-word';
			span.textContent = word;
			return span;
		}),
	);
}

/** True for a word span written by {@link writeLargeCaption}, which localization leaves alone. */
export function isCaptionWord(element: Element): boolean {
	return element.classList.contains('ribbon-caption-word');
}
