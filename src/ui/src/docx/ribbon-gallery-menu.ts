import { galleryLook, type GalleryKind } from './ribbon-gallery-look';
import { closeRibbonPopover, mountPopover } from './ribbon-popover';
import type { RibbonAction } from './ribbon-action';
import type { RibbonIcon } from './ribbon-icons';
import { menuSelect } from './ribbon-parts';
import { emit } from './events';
import { localeOf, translate } from './localization';

export { galleryLook } from './ribbon-gallery-look';

/** A command listed under a gallery's entries, such as Line Spacing Options. */
export interface GalleryCommand {
	label: string;
	action: RibbonAction;
}

/**
 * A dropdown command like `menuSelect`, but choosing opens a gallery of page thumbnails (Word's
 * Margins, Size, Orientation and Columns) instead of the browser's plain list. The real `<select>`
 * stays in the DOM so the editor keeps reading and writing `select.value`, and the keyboard still
 * changes it directly.
 */
export function menuGallery(
	label: string,
	icon: RibbonIcon,
	kind: GalleryKind,
	values: Array<[string, string]>,
	action: (value: string) => RibbonAction,
	options: { compact?: boolean; momentary?: boolean; commands?: GalleryCommand[] } = {},
): HTMLElement {
	const wrap = menuSelect(label, icon, values, action, options);
	wrap.classList.add('ribbon-menu-gallery');
	const control = wrap.querySelector('select')!;
	const open = () => openGallery(wrap, control, kind, options.commands ?? []);
	wrap.addEventListener('click', (event) => {
		if (control.disabled) return;
		event.preventDefault();
		open();
	});
	control.addEventListener('keydown', (event) => {
		if (event.key === 'Enter' || event.key === ' ' || (event.altKey && event.key === 'ArrowDown')) {
			event.preventDefault();
			open();
		}
	});
	return wrap;
}

function openGallery(
	anchor: HTMLElement,
	control: HTMLSelectElement,
	kind: GalleryKind,
	commands: GalleryCommand[],
): void {
	const pop = document.createElement('div');
	pop.className = 'ribbon-popover ribbon-gallery-menu';
	pop.setAttribute('role', 'menu');
	pop.setAttribute('aria-label', control.getAttribute('aria-label') ?? '');
	const locale = localeOf(anchor);
	const say = (text: string) => translate(locale, text as never);
	for (const option of [...control.options]) {
		const look = galleryLook(kind, option.value, say);
		const item = document.createElement('button');
		item.type = 'button';
		item.setAttribute('role', 'menuitemradio');
		item.setAttribute('aria-checked', String(option.value === control.value));
		item.dataset.value = option.value;
		const name = document.createElement('span');
		name.className = 'gallery-name';
		name.textContent = option.textContent;
		const words = document.createElement('span');
		words.className = 'gallery-words';
		words.append(name);
		if (look.detail) {
			const detail = document.createElement('span');
			detail.className = 'gallery-detail';
			detail.textContent = look.detail;
			words.append(detail);
		}
		item.append(look.thumb, words);
		item.addEventListener('mousedown', (event) => event.preventDefault());
		item.addEventListener('click', () => {
			closeRibbonPopover();
			control.value = option.value;
			control.dispatchEvent(new Event('change', { bubbles: true }));
		});
		pop.append(item);
	}
	for (const command of commands) {
		const item = document.createElement('button');
		item.type = 'button';
		item.setAttribute('role', 'menuitem');
		item.className = 'gallery-command';
		item.textContent = say(command.label);
		item.addEventListener('mousedown', (event) => event.preventDefault());
		item.addEventListener('click', () => {
			closeRibbonPopover();
			emit(anchor, 'ribbon-action', command.action);
		});
		pop.append(item);
	}
	mountPopover(anchor, pop, anchor);
	pop.querySelector<HTMLElement>('[aria-checked="true"]')?.focus({ preventScroll: true });
}
