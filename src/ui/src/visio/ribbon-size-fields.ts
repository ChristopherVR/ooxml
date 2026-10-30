/**
 * The Width and Height boxes of the Size group: inches of the selected shape, committed by
 * `ViewerInlineSize` (the same core resize the Size & Position pane makes).
 */
export function sizeFields(doc: Document): HTMLElement {
	const el = doc.createElement('div');
	el.className = 'ribbon-size-fields';
	for (const field of ['width', 'height'] as const) {
		const label = doc.createElement('label');
		const caption = doc.createElement('span');
		caption.textContent = field === 'width' ? 'Width' : 'Height';
		const input = doc.createElement('input');
		input.type = 'number';
		input.min = '0';
		input.step = 'any';
		input.disabled = true;
		input.dataset.inlineSize = field;
		input.setAttribute('aria-label', field === 'width' ? 'Width (in)' : 'Height (in)');
		label.append(caption, input);
		el.append(label);
	}
	return el;
}
