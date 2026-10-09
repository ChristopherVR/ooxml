import { buildStencilsView, masterList } from './shapes-sections';
import { STENCILS } from './stencil-catalog';
export {
	BASIC_SHAPES,
	findMaster,
	masterCreation,
	type Master,
	type MasterCreation,
} from './stencil-catalog';

/** Drag payload type for a stencil master; the value is the master id. */
export const MASTER_MIME = 'application/x-visio-viewer-master';

/**
 * Visio's Shapes window: Stencils and Search views, More Shapes and Quick Shapes rows and the
 * open stencils (`shapes-sections.ts`). Masters are dragged onto the page; Enter adds one at the
 * page centre.
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

	const stencils = doc.createElement('div');
	stencils.id = 'shapes-stencils';
	stencils.setAttribute('role', 'tabpanel');
	stencils.setAttribute('aria-labelledby', 'shapes-stencils-tab');
	const menus = buildStencilsView(doc, stencils, pane);

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
	// Search covers Basic Shapes and every built-in stencil, open or not, as Visio's does.
	const results = masterList(
		doc,
		STENCILS.flatMap((stencil) => stencil.masters),
	);
	const empty = doc.createElement('p');
	empty.className = 'shapes-empty';
	empty.textContent = 'No matching shapes in the built-in stencils.';
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
	pane.append(heading, views, stencils, search, ...menus);
	return pane;
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
