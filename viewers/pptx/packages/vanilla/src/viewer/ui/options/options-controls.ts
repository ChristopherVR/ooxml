import { createEl } from '../../render';

/**
 * The File > Options rows themselves are the shared options dialog's; this keeps the small
 * secondary action button the product panes (Customize Ribbon, Quick Access, clear cache) use.
 */

/** Append a small secondary action button (Reset / Clear cache / Add / Remove). */
export function appendOptionsAction(
	doc: Document,
	parent: HTMLElement,
	label: string,
	onClick: () => void,
): HTMLButtonElement {
	const button = createEl(doc, 'button', 'pptxv-options-action');
	button.type = 'button';
	button.textContent = label;
	button.addEventListener('click', onClick);
	parent.appendChild(button);
	return button;
}
