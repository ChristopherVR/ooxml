import { createIconSvg } from './icons.js';
import { parseSvgPreview } from './safe-svg.js';

/** One tile. `preview` is SVG markup, parsed as an SVG document with executable parts removed. */
export interface OfficeGalleryItem {
	id: string;
	label: string;
	applied?: boolean | undefined;
	preview?: string | undefined;
}

export interface OfficeGallerySection {
	title?: string | undefined;
	columns: number;
	tileWidth: number;
	tileHeight: number;
	items: readonly OfficeGalleryItem[];
}

/** A gallery, translated. With `command` it is a single button that picks the first item. */
export interface OfficeGalleryState {
	id: string;
	label: string;
	/** Trigger name in the inline strip; defaults to "More <label>". */
	moreLabel?: string | undefined;
	disabled?: boolean | undefined;
	command?: { icon: string; large?: boolean | undefined; hint?: string | undefined } | undefined;
	sections: readonly OfficeGallerySection[];
	/** Tiles shown inline (`mode="inline"`); defaults to the first section's items. */
	inline?: readonly OfficeGalleryItem[] | undefined;
}

export interface GalleryHooks {
	triggerAttribute: string;
	popupAttribute: string;
	itemAttribute: string;
	compactAttribute: string | null;
}

export interface GalleryPaint {
	state: OfficeGalleryState;
	disabled: boolean;
	inline: boolean;
	chevronOnly: boolean;
	open: boolean;
	icon: Element | null;
}

/** All gallery markup: the strip, the trigger and the popup. The element owns behaviour. */
export function createGalleryView(
	doc: Document,
	hooks: GalleryHooks,
	pick: (id: string) => void,
	toggle: () => void,
) {
	const root = doc.createElement('span');
	root.className = 'gallery-view';
	const strip = doc.createElement('div');
	strip.className = 'strip';
	const trigger = doc.createElement('button');
	trigger.type = 'button';
	trigger.className = 'trigger';
	if (hooks.compactAttribute) trigger.setAttribute(hooks.compactAttribute, '');
	trigger.setAttribute('aria-haspopup', 'dialog');
	trigger.addEventListener('click', toggle);
	const popup = doc.createElement('div');
	popup.className = 'popup';
	popup.hidden = true;
	popup.setAttribute('role', 'dialog');
	root.append(strip, trigger);
	const tile = (item: OfficeGalleryItem, w: number, h: number, disabled: boolean) => {
		const button = doc.createElement('button');
		button.type = 'button';
		button.className = 'tile';
		if (hooks.compactAttribute) button.setAttribute(hooks.compactAttribute, '');
		button.setAttribute(hooks.itemAttribute, item.id);
		button.disabled = disabled;
		button.setAttribute('aria-pressed', String(item.applied === true));
		button.setAttribute('aria-label', item.label);
		button.title = item.label;
		// Tile size is document data (the preview's own aspect), so it is set inline.
		button.style.width = `${w + 6}px`;
		button.style.height = `${h + 6}px`;
		const preview = item.preview ? parseSvgPreview(doc, item.preview) : null;
		if (preview) button.append(preview);
		button.addEventListener('click', () => pick(item.id));
		return button;
	};
	const label = (text: string) => {
		const span = doc.createElement('span');
		span.textContent = text;
		return span;
	};
	return {
		root,
		trigger,
		popup,
		clear() {
			strip.replaceChildren();
			popup.replaceChildren();
			trigger.removeAttribute(hooks.triggerAttribute);
			trigger.removeAttribute('aria-label');
			trigger.textContent = '';
		},
		setOpen(open: boolean) {
			popup.hidden = !open;
			if (open) root.append(popup);
			else popup.remove();
		},
		paint({ state, disabled, inline, chevronOnly, open, icon }: GalleryPaint) {
			const title = state.label;
			trigger.setAttribute(hooks.triggerAttribute, state.id);
			popup.setAttribute(hooks.popupAttribute, state.id);
			trigger.title = inline ? (state.moreLabel ?? `More ${title}`) : title;
			trigger.setAttribute('aria-label', trigger.title);
			trigger.disabled = disabled;
			trigger.replaceChildren();
			const command = state.command;
			trigger.classList.toggle('command', Boolean(command));
			trigger.classList.toggle('command-large', Boolean(command?.large));
			if (command) {
				// A one-button command: no panel, no chevron; the tooltip says why it is off.
				trigger.removeAttribute('aria-haspopup');
				trigger.removeAttribute('aria-expanded');
				trigger.title = command.hint ?? title;
				trigger.setAttribute('aria-label', title);
				trigger.append(icon ?? createIconSvg(doc), label(title));
			} else {
				trigger.setAttribute('aria-haspopup', 'dialog');
				if (!inline && !chevronOnly) trigger.append(icon ?? createIconSvg(doc), label(title));
				const chevron = doc.createElement('span');
				chevron.textContent = '⌄';
				chevron.setAttribute('aria-hidden', 'true');
				trigger.append(chevron);
			}
			strip.hidden = !inline;
			strip.style.display = inline ? '' : 'none';
			const first = state.sections[0];
			const inlineItems = state.inline ?? first?.items ?? [];
			strip.replaceChildren(
				...(inline
					? inlineItems.map((item) =>
							tile(item, first?.tileWidth ?? 40, first?.tileHeight ?? 30, disabled),
						)
					: []),
			);
			popup.setAttribute('aria-label', title);
			popup.replaceChildren(
				...(open ? state.sections : []).map((section) => {
					const container = doc.createElement('section');
					container.className = 'section';
					if (section.title) {
						const heading = doc.createElement('div');
						heading.className = 'heading';
						heading.textContent = section.title;
						container.append(heading);
					}
					const grid = doc.createElement('div');
					grid.className = 'grid';
					grid.style.gridTemplateColumns = `repeat(${section.columns}, minmax(0, 1fr))`;
					grid.append(
						...section.items.map((item) =>
							tile(item, section.tileWidth, section.tileHeight, disabled),
						),
					);
					container.append(grid);
					return container;
				}),
			);
		},
	};
}
