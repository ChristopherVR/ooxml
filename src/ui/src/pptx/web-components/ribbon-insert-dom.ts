import { RIBBON_ICON_PATHS } from './ribbon-icons';

const SVG_NS = 'http://www.w3.org/2000/svg';

/** Trusted shared artwork only; paths are never taken from host input. */
export function insertSvg(doc: Document, path: string, viewBox = '0 0 20 20'): SVGSVGElement {
	const svg = doc.createElementNS(SVG_NS, 'svg');
	svg.setAttribute('viewBox', viewBox);
	svg.setAttribute('aria-hidden', 'true');
	const shape = doc.createElementNS(SVG_NS, 'path');
	shape.setAttribute('d', path);
	svg.append(shape);
	return svg;
}

export function insertIcon(doc: Document, name: string): SVGSVGElement {
	return insertSvg(doc, RIBBON_ICON_PATHS[name] ?? '');
}

export interface InsertMenuItem {
	id: string;
	label: string;
	/** Optional glyph path, 24px grid unless {@link viewBox} says otherwise. */
	glyph?: string;
	viewBox?: string;
	/** CSS transform applied to the glyph (rotated or skewed shape presets). */
	transform?: string;
	/** Listed but not actionable: the editor has no such feature. */
	disabled?: boolean;
	/** Tooltip; the label by default. */
	title?: string;
	/** A heading drawn above the first entry of each section. */
	section?: string;
	/** Marks the entry that stands for the current choice (the active layout). */
	current?: boolean;
}

export interface InsertMenuOptions {
	/** Office "large" command: a 32px glyph over the label, chevron below. */
	large?: boolean;
	/** Lay the entries out as an icon gallery instead of a list. */
	grid?: boolean;
	/** Office split button: the face runs `primary`, the chevron below opens the menu. */
	split?: { primary: () => void };
}

export interface InsertMenuView {
	label: string;
	/** Tooltip of the face (the title and description). */
	title: string;
	/** Name of the chevron of a split button. */
	menuTitle?: string;
	items: readonly InsertMenuItem[];
	disabled: boolean;
	/** The split button's face can be disabled apart from its chevron. */
	primaryDisabled?: boolean;
	menuLabel: string;
}

export interface InsertMenu {
	el: HTMLElement;
	/** The control that opens the menu (the chevron of a split button). */
	trigger: HTMLButtonElement;
	sync(view: InsertMenuView): void;
	close(): void;
	isOpen(): boolean;
	place(): void;
}

/** Click, Enter/Space and arrow-key menu; replaces the hover-only per-binding popups. */
export function createInsertMenu(
	doc: Document,
	control: string,
	icon: string,
	pick: (id: string) => void,
	options: InsertMenuOptions = {},
): InsertMenu {
	const el = doc.createElement('div');
	el.className = options.split ? 'menu split' : 'menu';
	el.dataset.ribbonControl = control;
	const trigger = doc.createElement('button');
	trigger.type = 'button';
	trigger.className = options.large && !options.split ? 'trigger large' : 'trigger';
	trigger.dataset.pptxCompact = '';
	trigger.setAttribute('aria-haspopup', 'menu');
	trigger.setAttribute('aria-expanded', 'false');
	const text = doc.createElement('span');
	text.className = 'label';
	const primary = doc.createElement('button');
	if (options.split) {
		primary.type = 'button';
		primary.className = 'primary';
		primary.dataset.pptxCompact = '';
		primary.append(insertIcon(doc, icon), text);
		primary.addEventListener('click', options.split.primary);
		trigger.classList.add('caret');
		trigger.append(insertIcon(doc, 'chevronDown'));
	} else {
		trigger.append(insertIcon(doc, icon), text, insertIcon(doc, 'chevronDown'));
	}
	const list = doc.createElement('div');
	list.className = options.grid ? 'list grid' : 'list';
	list.setAttribute('role', 'menu');
	list.hidden = true;
	el.append(...(options.split ? [primary, trigger, list] : [trigger, list]));
	const items = () => [...list.querySelectorAll<HTMLButtonElement>('button:not(:disabled)')];
	const close = () => {
		list.hidden = true;
		trigger.setAttribute('aria-expanded', 'false');
	};
	const place = () => {
		if (list.hidden) {
			return;
		}
		const rect = el.getBoundingClientRect();
		const view = doc.defaultView;
		const box = list.getBoundingClientRect();
		const width = view?.innerWidth ?? 1024;
		const height = view?.innerHeight ?? 768;
		list.style.left = `${Math.max(8, Math.min(rect.left, width - box.width - 8))}px`;
		list.style.top = `${Math.max(8, Math.min(rect.bottom + 4, height - box.height - 8))}px`;
	};
	const open = (focusFirst: boolean) => {
		if (trigger.disabled) {
			return;
		}
		list.hidden = false;
		trigger.setAttribute('aria-expanded', 'true');
		place();
		if (focusFirst) {
			items()[0]?.focus();
		}
	};
	trigger.addEventListener('click', () => (list.hidden ? open(false) : close()));
	el.addEventListener('keydown', (event) => {
		const entries = items();
		const index = entries.indexOf(doc.activeElement as HTMLButtonElement);
		if (event.key === 'ArrowDown' || event.key === 'ArrowUp') {
			event.preventDefault();
			event.stopPropagation();
			if (list.hidden) {
				open(true);
			} else if (entries.length) {
				const step = event.key === 'ArrowDown' ? 1 : -1;
				entries[(index + step + entries.length) % entries.length].focus();
			}
		} else if (event.key === 'Tab') {
			close();
		} else if (event.key === ' ' || event.key === 'Enter') {
			// Keep native activation out of the viewer's slide-navigation handlers.
			event.stopPropagation();
		}
	});
	const entryNode = (item: InsertMenuItem) => {
		const button = doc.createElement('button');
		button.type = 'button';
		button.setAttribute('role', 'menuitem');
		button.dataset.pptxCompact = '';
		button.dataset.insertItem = item.id;
		if (item.disabled) {
			button.disabled = true;
		}
		if (item.current) {
			button.setAttribute('aria-current', 'true');
		}
		if (item.glyph) {
			const glyph = insertSvg(doc, item.glyph, item.viewBox ?? '0 0 24 24');
			if (item.transform && item.transform !== 'none') {
				glyph.style.transform = item.transform;
			}
			button.append(glyph);
		}
		if (options.grid) {
			button.title = item.title ?? item.label;
			button.setAttribute('aria-label', item.label);
		} else {
			button.title = item.title ?? '';
			button.append(doc.createTextNode(item.label));
		}
		button.addEventListener('click', () => {
			close();
			trigger.focus();
			pick(item.id);
		});
		return button;
	};
	return {
		el,
		trigger,
		close,
		place,
		isOpen: () => !list.hidden,
		sync(view) {
			text.textContent = view.label;
			if (options.split) {
				primary.title = view.title;
				primary.disabled = view.primaryDisabled ?? view.disabled;
				primary.setAttribute('aria-label', view.label);
				trigger.title = view.menuTitle ?? view.menuLabel;
				// The name stays apart from the face's, so each is one button for assistive tech.
				trigger.setAttribute('aria-label', view.menuLabel);
			} else {
				trigger.title = view.title;
			}
			trigger.disabled = view.disabled;
			list.setAttribute('aria-label', view.menuLabel);
			if (view.disabled) {
				close();
			}
			const key = view.items
				.map((item) => [item.id, item.label, item.disabled, item.title, item.section, item.current])
				.join('|');
			if (list.dataset.key !== key) {
				list.dataset.key = key;
				const nodes: HTMLElement[] = [];
				let section: string | undefined;
				for (const item of view.items) {
					if (item.section && item.section !== section) {
						const heading = doc.createElement('div');
						heading.className = 'heading';
						heading.setAttribute('role', 'presentation');
						heading.textContent = item.section;
						nodes.push(heading);
					}
					section = item.section;
					nodes.push(entryNode(item));
				}
				list.replaceChildren(...nodes);
			}
		},
	};
}
