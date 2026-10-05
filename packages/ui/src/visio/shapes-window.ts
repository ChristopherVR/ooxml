/** Drag payload type for a stencil master; the value is the master id. */
export const MASTER_MIME = 'application/x-visio-viewer-master';

interface Master {
	id: string;
	name: string;
	/** Preview outline in a 24x24 box (static application geometry). */
	path: string;
	/** Default size in inches when dropped. */
	size?: { width: number; height: number };
}
const MASTER_REASON = 'Needs core master drops; only Rectangle is available.';
/** Visio's Basic Shapes stencil, in Visio's order. */
export const BASIC_SHAPES: readonly Master[] = [
	{ id: 'rectangle', name: 'Rectangle', path: 'M3 6h18v12H3Z', size: { width: 1, height: 0.75 } },
	{ id: 'square', name: 'Square', path: 'M5 4h14v14H5Z' },
	{ id: 'ellipse', name: 'Ellipse', path: 'M2 12a10 6.5 0 1 0 20 0 10 6.5 0 1 0-20 0' },
	{ id: 'circle', name: 'Circle', path: 'M4 12a8 8 0 1 0 16 0 8 8 0 1 0-16 0' },
	{ id: 'triangle', name: 'Triangle', path: 'M12 4 21 20H3Z' },
	{ id: 'right-triangle', name: 'Right triangle', path: 'M4 4v16h16Z' },
	{ id: 'pentagon', name: 'Pentagon', path: 'M12 3l9 6.5-3.4 10.5H6.4L3 9.5Z' },
	{ id: 'hexagon', name: 'Hexagon', path: 'M7 4h10l5 8-5 8H7l-5-8Z' },
	{ id: 'octagon', name: 'Octagon', path: 'M8.5 3h7L21 8.5v7L15.5 21h-7L3 15.5v-7Z' },
	{
		id: 'star',
		name: '5-point star',
		path: 'm12 2.5 2.8 6.4 6.9.6-5.2 4.6 1.6 6.8L12 17.3l-6.1 3.6 1.6-6.8-5.2-4.6 6.9-.6Z',
	},
	{ id: 'diamond', name: 'Diamond', path: 'M12 2.5 21.5 12 12 21.5 2.5 12Z' },
	{
		id: 'rounded-rectangle',
		name: 'Rounded rectangle',
		path: 'M6 6h12a3 3 0 0 1 3 3v6a3 3 0 0 1-3 3H6a3 3 0 0 1-3-3V9a3 3 0 0 1 3-3Z',
	},
	{ id: 'cross', name: 'Cross', path: 'M9 3h6v6h6v6h-6v6H9v-6H3V9h6Z' },
	{ id: 'parallelogram', name: 'Parallelogram', path: 'M7 6h15l-5 12H2Z' },
	{ id: 'trapezoid', name: 'Trapezoid', path: 'M7 6h10l5 12H2Z' },
	{ id: 'can', name: 'Can', path: 'M5 6a7 2.5 0 1 0 14 0 7 2.5 0 1 0-14 0v12a7 2.5 0 0 0 14 0V6' },
	{ id: 'cube', name: 'Cube', path: 'M4 8h12v12H4ZM4 8l4-4h12l-4 4M20 4v12l-4 4' },
	{ id: 'chevron', name: 'Chevron', path: 'M3 5h13l5 7-5 7H3l5-7Z' },
];

const SVG_NS = 'http://www.w3.org/2000/svg';
function preview(doc: Document, path: string): SVGSVGElement {
	const svg = doc.createElementNS(SVG_NS, 'svg');
	svg.setAttribute('viewBox', '0 0 24 24');
	svg.setAttribute('aria-hidden', 'true');
	const outline = doc.createElementNS(SVG_NS, 'path');
	outline.setAttribute('d', path);
	svg.append(outline);
	return svg;
}

/**
 * Visio's Shapes window: Stencils and Search views, More Shapes and Quick Shapes rows and the
 * Basic Shapes stencil. Masters are dragged onto the page; Enter adds one at the page centre.
 */
export function createShapesWindow(doc: Document): HTMLElement {
	const pane = doc.createElement('aside');
	pane.id = 'shapes-pane';
	pane.className = 'shapes-pane';
	pane.setAttribute('aria-label', 'Shapes');
	const heading = doc.createElement('div');
	heading.className = 'pane-heading';
	const title = doc.createElement('span');
	title.textContent = 'Shapes';
	const collapse = doc.createElement('button');
	collapse.type = 'button';
	collapse.className = 'pane-collapse';
	collapse.dataset.chrome = 'shapes';
	collapse.setAttribute('aria-label', 'Close Shapes');
	collapse.title = 'Close Shapes';
	collapse.textContent = '‹';
	heading.append(title, collapse);

	const views = doc.createElement('div');
	views.className = 'shapes-views';
	views.setAttribute('role', 'tablist');
	views.setAttribute('aria-label', 'Shapes views');
	const view = (key: 'stencils' | 'search', label: string, selected: boolean) => {
		const tab = doc.createElement('button');
		tab.type = 'button';
		tab.id = `shapes-${key}-tab`;
		tab.dataset.shapesView = key;
		tab.textContent = label;
		tab.setAttribute('role', 'tab');
		tab.setAttribute('aria-selected', String(selected));
		tab.setAttribute('aria-controls', `shapes-${key}`);
		tab.tabIndex = selected ? 0 : -1;
		return tab;
	};
	views.append(view('stencils', 'Stencils', true), view('search', 'Search', false));

	const row = (label: string, reason: string) => {
		const button = doc.createElement('button');
		button.type = 'button';
		button.className = 'shapes-row';
		button.disabled = true;
		button.title = `${label}: not available yet. ${reason}`;
		button.textContent = label;
		return button;
	};
	const stencilTitle = doc.createElement('div');
	stencilTitle.className = 'stencil-title';
	stencilTitle.textContent = 'Basic Shapes';
	const stencils = doc.createElement('div');
	stencils.id = 'shapes-stencils';
	stencils.setAttribute('role', 'tabpanel');
	stencils.setAttribute('aria-labelledby', 'shapes-stencils-tab');
	stencils.append(
		row('More Shapes', 'Needs stencil files.'),
		row('Quick Shapes', 'Needs stencil files.'),
		stencilTitle,
		masterList(doc, BASIC_SHAPES),
	);

	const search = doc.createElement('div');
	search.id = 'shapes-search';
	search.hidden = true;
	search.setAttribute('role', 'tabpanel');
	search.setAttribute('aria-labelledby', 'shapes-search-tab');
	const field = doc.createElement('input');
	field.type = 'search';
	field.className = 'shapes-search-field';
	field.placeholder = 'Search for shapes';
	field.setAttribute('aria-label', 'Search for shapes');
	field.autocomplete = 'off';
	const results = masterList(doc, BASIC_SHAPES);
	const empty = doc.createElement('p');
	empty.className = 'shapes-empty';
	empty.textContent = 'No matching shapes in Basic Shapes.';
	empty.hidden = true;
	field.addEventListener('input', () => {
		const query = field.value.trim().toLowerCase();
		let shown = 0;
		for (const item of results.querySelectorAll<HTMLElement>('li')) {
			item.hidden = !!query && !item.dataset.name!.toLowerCase().includes(query);
			if (!item.hidden) shown++;
		}
		empty.hidden = shown > 0;
	});
	search.append(field, results, empty);

	views.addEventListener('click', (event) => {
		const tab = (event.target as Element).closest<HTMLButtonElement>('[data-shapes-view]');
		if (!tab) return;
		for (const candidate of views.querySelectorAll<HTMLButtonElement>('[data-shapes-view]')) {
			const selected = candidate === tab;
			candidate.setAttribute('aria-selected', String(selected));
			candidate.tabIndex = selected ? 0 : -1;
		}
		stencils.hidden = tab.dataset.shapesView !== 'stencils';
		search.hidden = !stencils.hidden;
		if (!search.hidden) field.focus();
	});
	pane.append(heading, views, stencils, search);
	return pane;
}

function masterList(doc: Document, masters: readonly Master[]): HTMLUListElement {
	const list = doc.createElement('ul');
	list.className = 'masters';
	for (const master of masters) {
		const item = doc.createElement('li');
		item.dataset.name = master.name;
		const button = doc.createElement('button');
		button.type = 'button';
		button.className = 'master';
		button.dataset.master = master.id;
		const label = doc.createElement('span');
		label.textContent = master.name;
		button.append(preview(doc, master.path), label);
		if (master.size) {
			button.draggable = true;
			button.title = `${master.name}: drag onto the page, or press Enter to add it at the centre.`;
		} else {
			button.setAttribute('aria-disabled', 'true');
			button.dataset.unsupported = '';
			button.title = `${master.name}: not available yet. ${MASTER_REASON}`;
		}
		item.append(button);
		list.append(item);
	}
	return list;
}

/** Default size of a stencil master the viewer can create, or nothing when unsupported. */
export function masterSize(id: string): { width: number; height: number } | undefined {
	return BASIC_SHAPES.find((master) => master.id === id)?.size;
}

/** Visio's minimised Shapes window: a narrow strip that reopens the window. */
export function createShapesStrip(doc: Document): HTMLButtonElement {
	const strip = doc.createElement('button');
	strip.type = 'button';
	strip.className = 'shapes-strip';
	strip.dataset.chrome = 'shapes';
	strip.hidden = true;
	strip.setAttribute('aria-label', 'Open Shapes');
	strip.title = 'Open Shapes';
	const arrow = doc.createElement('span');
	arrow.setAttribute('aria-hidden', 'true');
	arrow.textContent = '›';
	const label = doc.createElement('span');
	label.className = 'shapes-strip-label';
	label.textContent = 'Shapes';
	strip.append(arrow, label);
	return strip;
}
