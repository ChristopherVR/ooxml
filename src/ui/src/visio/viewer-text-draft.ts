import type { VisioDrawingBounds } from 'ooxml-core/visio/ui';

/** Plain browser input, separate from the saved SVG text and its layout. */
export function createTextDraft(
	viewport: HTMLElement,
	svg: SVGSVGElement,
	bounds: VisioDrawingBounds,
	apply: () => void,
	cancel: () => void,
) {
	const doc = viewport.ownerDocument;
	const element = doc.createElement('section');
	element.className = 'text-box-draft';
	element.dataset.textDraft = '';
	element.setAttribute('aria-label', 'New text box');
	const label = doc.createElement('label');
	label.textContent = 'Text box text';
	const input = doc.createElement('textarea');
	input.rows = 4;
	input.setAttribute('aria-label', 'Text box text');
	label.append(input);
	const help = doc.createElement('p');
	help.textContent =
		'Add plain text within the drawn bounds. Ctrl+Enter adds the text box; Escape cancels.';
	const error = doc.createElement('p');
	error.setAttribute('role', 'alert');
	error.hidden = true;
	const actions = doc.createElement('div');
	const buttons = ['Add text box', 'Cancel'].map((name) => {
		const button = doc.createElement('office-ui-button') as HTMLElement & { disabled: boolean };
		button.setAttribute('label', name);
		button.setAttribute('command', name === 'Cancel' ? 'cancel-text-box' : 'apply-text-box');
		button.addEventListener('office-command', name === 'Cancel' ? cancel : apply);
		actions.append(button);
		return button;
	});
	input.addEventListener('keydown', (event) => {
		event.stopPropagation();
		if (event.isComposing) return;
		if (event.key === 'Escape') {
			event.preventDefault();
			event.stopPropagation();
			cancel();
		} else if (event.key === 'Enter' && (event.ctrlKey || event.metaKey)) {
			event.preventDefault();
			event.stopPropagation();
			apply();
		}
	});
	for (const name of ['click', 'pointerdown', 'contextmenu'])
		element.addEventListener(name, (event) => event.stopPropagation());
	element.append(label, help, error, actions);
	const matrix = svg.getScreenCTM();
	const box = viewport.getBoundingClientRect();
	let left = viewport.scrollLeft,
		top = viewport.scrollTop;
	if (matrix) {
		const point = new DOMPoint(bounds.x, bounds.y).matrixTransform(matrix);
		left += point.x - box.left;
		top += point.y - box.top;
	}
	viewport.append(element);
	const width = viewport.clientWidth || box.width || doc.defaultView?.innerWidth || 320;
	const height = viewport.clientHeight || box.height || doc.defaultView?.innerHeight || 240;
	element.style.maxHeight = `${Math.max(0, height - 16)}px`;
	element.style.left = `${viewport.scrollLeft + Math.max(8, Math.min(left - viewport.scrollLeft, width - element.offsetWidth - 8))}px`;
	element.style.top = `${viewport.scrollTop + Math.max(8, Math.min(top - viewport.scrollTop, height - element.offsetHeight - 8))}px`;
	input.focus();
	return {
		element,
		input,
		error,
		setBusy(busy: boolean) {
			element.setAttribute('aria-busy', String(busy));
			input.disabled = busy;
			buttons[0]!.disabled = busy;
		},
		dispose() {
			element.remove();
		},
	};
}
