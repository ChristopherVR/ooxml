import { createIconSvg, paintIcon } from './icons.js';

/** A text segment at the start of the bar (page count, words, language, save state). */
export interface OfficeStatusText {
	id: string;
	text: string;
	title?: string | undefined;
	tone?: 'idle' | 'saving' | 'error' | undefined;
	/** Announce changes politely (a page counter). */
	live?: boolean | undefined;
	/** Hide on phone widths. */
	narrowHide?: boolean | undefined;
}

/** A status-bar button (Notes, Comments, a view switch). `label` is its name and ScreenTip. */
export interface OfficeStatusButton {
	id: string;
	icon: string;
	label: string;
	/** Visible text after the glyph; hidden on phone widths. */
	text?: string | undefined;
	pressed?: boolean | undefined;
	hidden?: boolean | undefined;
}

/** Everything the controlled bar shows, already translated. Absent parts are not rendered. */
export interface OfficeStatusBarState {
	items?: readonly OfficeStatusText[] | undefined;
	toggles?: readonly OfficeStatusButton[] | undefined;
	/** The view switches (Normal, Sorter, Reading); hidden while empty or `false`. */
	views?: readonly OfficeStatusButton[] | false | undefined;
	/** Zoom out, percentage (zoom to fit) and zoom in; emits `zoomOut`, `zoomFit`, `zoomIn`. */
	zoom?:
		| {
				percent: number;
				outLabel: string;
				fitLabel: string;
				inLabel: string;
				/** Glyphs; default `minus` and `plus`. */
				outIcon?: string | undefined;
				inIcon?: string | undefined;
		  }
		| undefined;
}

export interface StatusBarView {
	bar: HTMLElement;
	render(state: OfficeStatusBarState | undefined): void;
}

/**
 * The controlled status bar: built once, then patched in place so a focused button survives
 * updates. Slots stay in the row: the default slot after the texts, `collaboration` between the
 * views and the zoom, and `end` before the zoom.
 */
export function createStatusBarView(doc: Document, activate: (id: string) => void): StatusBarView {
	const el = (tag: string, className: string): HTMLElement => {
		const node = doc.createElement(tag);
		node.className = className;
		return node;
	};
	const sep = (className: string): HTMLElement => {
		const node = el('i', className);
		node.setAttribute('aria-hidden', 'true');
		return node;
	};
	const slot = (name: string): HTMLSlotElement => {
		const node = doc.createElement('slot');
		if (name) node.name = name;
		return node;
	};
	const button = (id: string, className: string): HTMLButtonElement => {
		const node = el('button', className) as HTMLButtonElement;
		node.type = 'button';
		node.dataset.id = id;
		node.addEventListener('click', () => activate(id));
		// Keep native activation out of the host document's navigation shortcuts.
		node.addEventListener('keydown', (event) => {
			const { key, ctrlKey, metaKey, altKey } = event;
			if ((key === ' ' || key === 'Enter') && !ctrlKey && !metaKey && !altKey)
				event.stopPropagation();
		});
		return node;
	};
	const paint = (node: HTMLButtonElement, spec: OfficeStatusButton): void => {
		let svg = node.querySelector('svg');
		if (!svg) svg = node.insertBefore(createIconSvg(doc), node.firstChild);
		paintIcon(svg, spec.icon);
		node.title = spec.label;
		node.setAttribute('aria-label', spec.label);
		if (spec.pressed !== undefined) node.setAttribute('aria-pressed', String(spec.pressed));
		else node.removeAttribute('aria-pressed');
		node.hidden = spec.hidden === true;
		let text = node.querySelector<HTMLElement>('.label');
		if (spec.text) {
			text ??= node.appendChild(el('span', 'label'));
			text.textContent = spec.text;
		} else {
			text?.remove();
		}
	};
	/** Keyed children patched in place; reordered only when the order changes. */
	const keyed = <T extends { id: string }>(
		parent: HTMLElement,
		make: (spec: T) => HTMLElement,
		patch: (node: HTMLElement, spec: T) => void,
	) => {
		const nodes = new Map<string, HTMLElement>();
		return (specs: readonly T[]): void => {
			for (const [id, node] of nodes)
				if (!specs.some((spec) => spec.id === id)) {
					node.remove();
					nodes.delete(id);
				}
			for (const spec of specs) {
				let node = nodes.get(spec.id);
				if (!node) {
					node = make(spec);
					nodes.set(spec.id, node);
				}
				patch(node, spec);
			}
			const current = [...parent.children];
			if (specs.some((spec, index) => current[index] !== nodes.get(spec.id)))
				parent.append(...specs.map((spec) => nodes.get(spec.id)!));
		};
	};

	const items = el('span', 'items');
	const renderItems = keyed<OfficeStatusText>(
		items,
		(spec) => {
			const node = el('span', 'item');
			node.dataset.item = spec.id;
			return node;
		},
		(node, spec) => {
			node.textContent = spec.text;
			if (spec.title) node.title = spec.title;
			else node.removeAttribute('title');
			node.classList.toggle('saving', spec.tone === 'saving');
			node.classList.toggle('error', spec.tone === 'error');
			node.classList.toggle('narrow-hide', spec.narrowHide === true);
			if (spec.live) node.setAttribute('aria-live', 'polite');
			else node.removeAttribute('aria-live');
		},
	);
	const spacer = el('span', 'spacer');
	const toggles = el('span', 'toggles');
	const renderToggles = keyed<OfficeStatusButton>(
		toggles,
		(spec) => button(spec.id, 'toggle'),
		(node, spec) => paint(node as HTMLButtonElement, spec),
	);
	const viewSep = sep('sep tight');
	const views = el('div', 'group views');
	const renderViews = keyed<OfficeStatusButton>(
		views,
		(spec) => button(spec.id, 'view'),
		(node, spec) => paint(node as HTMLButtonElement, spec),
	);
	const collaboration = slot('collaboration');
	const collabSep = sep('sep tight');
	collabSep.hidden = true;
	collaboration.addEventListener('slotchange', () => {
		collabSep.hidden = collaboration.assignedNodes().length === 0;
	});
	const zoomSep = sep('sep tight');
	const zoom = el('div', 'group zoom');
	const zoomOut = button('zoomOut', 'zoom-step');
	zoomOut.append(createIconSvg(doc));
	const zoomFit = button('zoomFit', 'zoom-fit');
	const zoomIn = button('zoomIn', 'zoom-step');
	zoomIn.append(createIconSvg(doc));
	zoom.append(zoomOut, zoomFit, zoomIn);
	const bar = el('div', 'bar');
	bar.setAttribute('part', 'bar');
	bar.append(
		items,
		slot(''),
		spacer,
		toggles,
		viewSep,
		views,
		collabSep,
		collaboration,
		slot('end'),
		zoomSep,
		zoom,
	);
	const name = (node: HTMLButtonElement, label: string): void => {
		node.title = label;
		node.setAttribute('aria-label', label);
	};

	return {
		bar,
		render(state) {
			const itemSpecs = state?.items ?? [];
			renderItems(itemSpecs);
			items.hidden = itemSpecs.length === 0;
			const toggleSpecs = state?.toggles ?? [];
			renderToggles(toggleSpecs);
			toggles.hidden = toggleSpecs.length === 0;
			const viewSpecs = state?.views || [];
			renderViews(viewSpecs);
			views.hidden = viewSpecs.length === 0;
			viewSep.hidden = views.hidden || toggles.hidden;
			spacer.hidden = !state;
			const z = state?.zoom;
			zoomSep.hidden = zoom.hidden = !z;
			if (z) {
				name(zoomOut, z.outLabel);
				name(zoomFit, z.fitLabel);
				name(zoomIn, z.inLabel);
				paintIcon(zoomOut.querySelector('svg')!, z.outIcon ?? 'minus');
				paintIcon(zoomIn.querySelector('svg')!, z.inIcon ?? 'plus');
				zoomFit.textContent = `${Math.round(z.percent)}%`;
			}
		},
	};
}
