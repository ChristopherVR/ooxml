import { buildStencilsView, type StencilsView } from './shapes-sections';
import { masterList } from './shapes-master-list';
import { rememberMasterFocus } from './shapes-focus';
import { STENCILS } from './stencil-catalog';
import type { ShapesDocument } from './shapes-document';
import type { ViewerController } from './controller';
import { ribbonLocale } from './ribbon-localize';
import { normalizeVisioLocale } from './ribbon-locale';
import { stencilText, translateStencilText } from './stencil-locale';

const views = new WeakMap<HTMLElement, Pick<StencilsView, 'setDocument' | 'refreshLocale'>>();
const controllers = new WeakMap<HTMLElement, ViewerController>();

export function bindShapesController(pane: HTMLElement, controller?: ViewerController): void {
	if (controller) controllers.set(pane, controller);
	else controllers.delete(pane);
}
/**
 * Show a drawing in a Shapes window: its Document Stencil, the stencils it docks, and its masters
 * in Search shapes. Cheap to call again for the same drawing. `opened` is false when the same
 * drawing was only edited, so the stencils the user unfolded stay as they are.
 */
export function setShapesDocument(pane: HTMLElement, next: ShapesDocument, opened = true): void {
	views.get(pane)?.setDocument(next, opened);
}

export function setShapesLocale(pane: HTMLElement, locale: string): void {
	const previous = ribbonLocale(pane);
	pane.dataset.visioLocale = locale;
	if (ribbonLocale(pane) !== previous) views.get(pane)?.refreshLocale();
}
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
 * Visio's Shapes window: the Search shapes box, More Shapes and Quick Shapes rows and the open
 * stencils (`shapes-sections.ts`). Masters are dragged onto the page; Enter adds one at the
 * page centre.
 */
export function createShapesWindow(doc: Document): HTMLElement {
	const pane = doc.createElement('aside');
	pane.id = 'shapes-pane';
	pane.className = 'shapes-pane';
	const heading = doc.createElement('div');
	heading.className = 'pane-heading';
	const title = doc.createElement('span');
	const collapse = doc.createElement('button');
	collapse.type = 'button';
	collapse.className = 'pane-collapse';
	collapse.dataset.chrome = 'shapes';
	collapse.textContent = '‹';
	heading.append(title, collapse);

	// Visio's Search shapes box sits under the title; typing swaps the stencils for the results.
	const box = doc.createElement('div');
	box.className = 'shapes-search-box';
	const field = doc.createElement('input');
	field.type = 'search';
	field.className = 'shapes-search-field';
	field.setAttribute('aria-controls', 'shapes-search');
	field.autocomplete = 'off';
	const glass = doc.createElement('office-ui-icon');
	glass.setAttribute('name', 'search');
	glass.setAttribute('aria-hidden', 'true');
	box.append(field, glass);

	const stencils = doc.createElement('div');
	stencils.id = 'shapes-stencils';
	const view = buildStencilsView(doc, stencils, pane, () => controllers.get(pane));

	const search = doc.createElement('div');
	search.id = 'shapes-search';
	search.hidden = true;
	search.setAttribute('role', 'region');
	// Search covers the drawing's masters, Basic Shapes and every built-in stencil, open or not.
	const builtIn = STENCILS.flatMap((stencil) => stencil.masters);
	let masters = builtIn;
	const t = (text: string) => translateStencilText(pane, text);
	const results = masterList(doc, builtIn, t);
	const empty = doc.createElement('p');
	empty.className = 'shapes-empty';
	empty.hidden = true;
	const filter = () => {
		const query = field.value.trim().toLowerCase();
		let shown = 0;
		for (const item of results.querySelectorAll<HTMLElement>('li')) {
			item.hidden = !item.dataset.name!.toLowerCase().includes(query);
			if (!item.hidden) shown++;
		}
		empty.hidden = shown > 0;
		search.hidden = !query;
		stencils.hidden = !!query;
	};
	field.addEventListener('input', filter);
	search.append(results, empty);
	pane.append(heading, box, stencils, search, ...view.menus);
	const localize = () => {
		title.textContent = t('Shapes');
		pane.setAttribute('aria-label', title.textContent);
		collapse.title = t('Close Shapes');
		collapse.setAttribute('aria-label', collapse.title);
		field.placeholder = t('Search shapes');
		field.setAttribute('aria-label', field.placeholder);
		search.setAttribute('aria-label', t('Search results'));
		empty.textContent = t('No shapes match your search.');
	};
	const renderResults = () => {
		const restoreFocus = rememberMasterFocus(search, field);
		results.replaceChildren(...masterList(doc, masters, t).children);
		filter();
		restoreFocus?.();
	};
	localize();
	let key = '';
	views.set(pane, {
		refreshLocale() {
			localize();
			view.refreshLocale();
			renderResults();
		},
		setDocument(next, opened) {
			view.setDocument(next, opened);
			if (next.key === key) return;
			key = next.key;
			masters = [...next.masters, ...builtIn];
			renderResults();
		},
	});
	return pane;
}

/** Visio's minimised Shapes window: a narrow strip that reopens the window. */
export function createShapesStrip(doc: Document): HTMLButtonElement {
	const strip = doc.createElement('button');
	strip.type = 'button';
	strip.className = 'shapes-strip';
	strip.dataset.chrome = 'shapes';
	strip.hidden = true;
	const arrow = doc.createElement('span');
	arrow.setAttribute('aria-hidden', 'true');
	arrow.textContent = '›';
	const label = doc.createElement('span');
	label.className = 'shapes-strip-label';
	strip.append(arrow, label);
	setShapesStripLocale(strip, 'en');
	return strip;
}

export function setShapesStripLocale(strip: HTMLButtonElement, locale: string): void {
	const language = normalizeVisioLocale(locale);
	strip.title = stencilText(language, 'Open Shapes');
	strip.setAttribute('aria-label', strip.title);
	const label = strip.querySelector('.shapes-strip-label');
	if (label) label.textContent = stencilText(language, 'Shapes');
}
