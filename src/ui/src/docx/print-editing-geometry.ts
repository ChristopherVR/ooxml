export function glyphRect(element: HTMLElement, from: number, to = from): DOMRect {
	const start = Number(element.dataset.sourceStart);
	const end = Number(element.dataset.sourceEnd);
	const text = element.firstChild;
	const box = element.getBoundingClientRect();
	if (text?.nodeType === Node.TEXT_NODE && end > start) {
		const length = text.textContent?.length ?? 0;
		const offset = (value: number) =>
			Math.max(0, Math.min(length, Math.round(((value - start) * length) / (end - start))));
		const range = element.ownerDocument.createRange();
		range.setStart(text, offset(from));
		range.setEnd(text, offset(to));
		if (typeof range.getBoundingClientRect === 'function') {
			const measured = range.getBoundingClientRect();
			if (measured.height) return measured;
		}
	}
	const ratio = end > start ? Math.max(0, Math.min(1, (from - start) / (end - start))) : 0;
	const width =
		end > start ? box.width * Math.max(0, Math.min(1, (to - from) / (end - start))) : box.width;
	return new DOMRect(box.left + box.width * ratio, box.top, from === to ? 0 : width, box.height);
}

export function caretRect(line: HTMLElement, offset: number): DOMRect {
	const fragments = [...line.querySelectorAll<HTMLElement>('[data-source-start]')];
	const fragment =
		fragments.find(
			(el) => Number(el.dataset.sourceStart) <= offset && Number(el.dataset.sourceEnd) >= offset,
		) ?? fragments.at(-1);
	const rect = fragment ? glyphRect(fragment, offset) : line.getBoundingClientRect();
	const lineRect = line.getBoundingClientRect();
	return new DOMRect(fragment ? rect.left : lineRect.left, lineRect.top, 0, lineRect.height);
}
