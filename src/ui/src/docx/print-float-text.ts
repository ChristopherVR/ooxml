import type { EditorView } from 'prosemirror-view';
import { setShapeText } from './shape-commands';

/**
 * Edits the text of a floating shape or text box in place on its page: a textarea laid over the
 * shape, one paragraph per line. Enter adds a line, Escape cancels, Ctrl+Enter or leaving the
 * field applies it as one transaction (one undo step). Returns a function that applies and closes the editor
 * ; undefined when the shape's text cannot be edited.
 */
export function editShapeText(
	view: EditorView,
	element: HTMLElement,
	lines: string[],
	onClose: () => void,
): (() => void) | undefined {
	const sheet = element.closest<HTMLElement>('.dve-print-page');
	if (!sheet || !view.editable) return undefined;
	const field = document.createElement('textarea');
	field.className = 'dve-print-float-text';
	field.setAttribute('aria-label', 'Shape text');
	field.value = lines.join('\n');
	Object.assign(field.style, {
		left: element.style.left,
		top: element.style.top,
		width: element.style.width,
		height: element.style.height,
	});
	let open = true;
	const close = (apply: boolean) => {
		if (!open) return;
		open = false;
		const value = field.value;
		field.remove();
		if (apply && value !== lines.join('\n')) setShapeText(view, value.split(/\r\n|\r|\n/));
		onClose();
		if (!view.isDestroyed) view.focus();
	};
	field.addEventListener('keydown', (event) => {
		event.stopPropagation();
		if (event.key === 'Escape') {
			event.preventDefault();
			close(false);
		} else if (event.key === 'Enter' && (event.ctrlKey || event.metaKey)) {
			event.preventDefault();
			close(true);
		}
	});
	field.addEventListener('pointerdown', (event) => event.stopPropagation());
	field.addEventListener('blur', () => close(true));
	sheet.append(field);
	field.focus();
	field.setSelectionRange(field.value.length, field.value.length);
	return () => close(true);
}
