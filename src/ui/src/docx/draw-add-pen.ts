// Word's Add Pen: a large dropdown command (icon, caption, caret) like Excel's, whose menu adds a
// pen, a pencil or a highlighter to the gallery. Keyboard: Enter, Space or ArrowDown open the menu,
// arrows move through it, Escape closes it and gives focus back to the command.
import { emit } from './events';
import { localeOf, translate, type LocalizationKey } from './localization';
import { mountPopover, closeRibbonPopover } from './ribbon-popover';
import { ribbonIcon } from './ribbon-icons';

const KINDS = [
	['pen', 'Pen'],
	['pencil', 'Pencil'],
	['highlighter', 'Highlighter'],
] as const;

function openMenu(button: HTMLButtonElement): void {
	const locale = localeOf(button);
	const menu = document.createElement('div');
	menu.className = 'ribbon-popover dve-add-pen-menu';
	menu.setAttribute('role', 'menu');
	menu.setAttribute('aria-label', translate(locale, 'Add Pen' as LocalizationKey));
	const items = KINDS.map(([value, label]) => {
		const item = document.createElement('button');
		item.type = 'button';
		item.className = 'dve-add-pen-item';
		item.setAttribute('role', 'menuitem');
		item.tabIndex = -1;
		item.textContent = translate(locale, label as LocalizationKey);
		item.addEventListener('click', () => {
			closeRibbonPopover();
			button.focus();
			emit(button, 'ribbon-action', { type: 'draw', key: 'addPen', value });
		});
		item.addEventListener('keydown', (event) => {
			const at = items.indexOf(item);
			const step = event.key === 'ArrowDown' ? 1 : event.key === 'ArrowUp' ? -1 : 0;
			const next = event.key === 'Home' ? 0 : event.key === 'End' ? items.length - 1 : at + step;
			if (!step && event.key !== 'Home' && event.key !== 'End') return;
			event.preventDefault();
			event.stopPropagation();
			items[(next + items.length) % items.length]?.focus();
		});
		menu.append(item);
		return item;
	});
	if (mountPopover(button, menu, button)) items[0]?.focus();
}

/** The Add Pen command, marked `data-draw="addPen"` for the tab's state. */
export function addPenButton(): HTMLButtonElement {
	const button = document.createElement('button');
	button.type = 'button';
	button.className = 'ribbon-large dve-add-pen';
	button.dataset.draw = 'addPen';
	button.dataset.localearialabel = 'Add Pen';
	button.setAttribute('aria-label', 'Add Pen');
	button.setAttribute('aria-haspopup', 'menu');
	button.setAttribute('aria-expanded', 'false');
	button.title = 'Add Pen';
	const caption = document.createElement('span');
	caption.textContent = 'Add Pen';
	button.append(ribbonIcon('addPen', 28), caption, ribbonIcon('caret', 12));
	button.addEventListener('click', () => {
		if (!button.disabled) openMenu(button);
	});
	button.addEventListener('keydown', (event) => {
		if (event.key !== 'ArrowDown' || button.disabled) return;
		event.preventDefault();
		if (button.getAttribute('aria-expanded') !== 'true') openMenu(button);
	});
	return button;
}
