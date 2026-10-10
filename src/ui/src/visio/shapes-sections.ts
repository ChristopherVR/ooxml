import {
	BASIC_STENCIL_ID,
	STENCILS,
	findMaster,
	type Master,
	type Stencil,
} from './stencil-catalog';

/** Per-viewer Shapes window preferences: opened stencils and each stencil's Quick Shapes. */
export const SHAPES_STORAGE_KEY = 'ooxml-ui.visio.shapes';
/** Visio shows a stencil's first masters as its Quick Shapes until the user changes them. */
const DEFAULT_QUICK = 4;
interface Saved {
	open: string[];
	quick: Record<string, string[]>;
}
/** Undefined until the shared menu element is registered. */
type Menu = HTMLElement & { openAt?(x: number, y: number): void };

function storage(doc: Document): Storage | undefined {
	try {
		return doc.defaultView?.localStorage ?? undefined;
	} catch {
		return undefined;
	}
}
function load(doc: Document): Saved {
	const saved: Saved = { open: [], quick: {} };
	try {
		const raw = storage(doc)?.getItem(SHAPES_STORAGE_KEY);
		const value = raw ? (JSON.parse(raw) as Partial<Saved>) : {};
		const known = new Set(STENCILS.map((stencil) => stencil.id));
		if (Array.isArray(value.open))
			saved.open = value.open.filter(
				(id): id is string => typeof id === 'string' && known.has(id) && id !== BASIC_STENCIL_ID,
			);
		for (const [stencil, ids] of Object.entries(value.quick ?? {}))
			if (known.has(stencil) && Array.isArray(ids))
				saved.quick[stencil] = ids.filter(
					(id): id is string => typeof id === 'string' && findMaster(id)?.stencil.id === stencil,
				);
	} catch {
		// Blocked or corrupt storage: start from Visio's defaults.
	}
	return saved;
}
function save(doc: Document, saved: Saved): void {
	try {
		storage(doc)?.setItem(SHAPES_STORAGE_KEY, JSON.stringify(saved));
	} catch {
		// Preferences are a convenience; the window works without them.
	}
}

/** One master button: drag it onto the page, or activate it to add it at the page centre. */
export function masterButton(doc: Document, master: Master): HTMLLIElement {
	const item = doc.createElement('li');
	item.dataset.name = master.name;
	const button = doc.createElement('button');
	button.type = 'button';
	button.className = 'master';
	button.dataset.master = master.id;
	const svg = doc.createElementNS('http://www.w3.org/2000/svg', 'svg');
	svg.setAttribute('viewBox', '0 0 24 24');
	svg.setAttribute('aria-hidden', 'true');
	const outline = doc.createElementNS('http://www.w3.org/2000/svg', 'path');
	outline.setAttribute('d', master.path);
	svg.append(outline);
	const label = doc.createElement('span');
	label.textContent = master.name;
	button.append(svg, label);
	button.draggable = true;
	button.title = `${master.name}: drag onto the page, or press Enter to add it at the centre.`;
	item.append(button);
	return item;
}
export function masterList(doc: Document, masters: readonly Master[]): HTMLUListElement {
	const list = doc.createElement('ul');
	list.className = 'masters';
	list.append(...masters.map((master) => masterButton(doc, master)));
	return list;
}
function menuItem(doc: Document, command: string, label: string, checked?: boolean): HTMLElement {
	const item = doc.createElement('office-ui-menu-item');
	item.setAttribute('command', command);
	item.setAttribute('label', label);
	if (checked !== undefined) item.setAttribute('checked', String(checked));
	return item;
}

function unsupportedItem(item: HTMLElement, reason: string): HTMLElement {
	item.setAttribute('disabled', '');
	item.dataset.unsupported = '';
	item.setAttribute('title', reason);
	return item;
}

/**
 * The Stencils view below the view tabs: More Shapes (a menu of the built-in stencils), Quick
 * Shapes (favourites across the open stencils) and one collapsible section per open stencil,
 * Basic Shapes first. Opened stencils and Quick Shapes persist per viewer in localStorage.
 */
export function buildStencilsView(
	doc: Document,
	panel: HTMLElement,
	pane: HTMLElement,
): HTMLElement[] {
	const saved = load(doc);
	const row = (id: string, label: string) => {
		const button = doc.createElement('button');
		button.type = 'button';
		button.id = id;
		button.className = 'shapes-row';
		button.textContent = label;
		return button;
	};
	const more = row('shapes-more', 'More Shapes');
	more.setAttribute('aria-haspopup', 'menu');
	more.title = 'More Shapes: open another stencil.';
	const quickToggle = row('shapes-quick-toggle', 'Quick Shapes');
	quickToggle.setAttribute('aria-expanded', 'false');
	quickToggle.setAttribute('aria-controls', 'shapes-quick');
	quickToggle.title = 'Quick Shapes: favourite masters of every open stencil.';
	const quick = doc.createElement('div');
	quick.id = 'shapes-quick';
	quick.hidden = true;
	const sections = doc.createElement('div');
	sections.id = 'shapes-sections';

	const moreMenu = doc.createElement('office-ui-context-menu') as Menu;
	moreMenu.dataset.shapesMenu = 'more';
	moreMenu.setAttribute('label', 'More Shapes');
	const masterMenu = doc.createElement('office-ui-context-menu') as Menu;
	masterMenu.dataset.shapesMenu = 'master';
	masterMenu.setAttribute('label', 'Master');
	let menuMaster: string | undefined;

	const openIds = () => [BASIC_STENCIL_ID, ...saved.open];
	const quickIds = (stencil: Stencil) =>
		saved.quick[stencil.id] ?? stencil.masters.slice(0, DEFAULT_QUICK).map((master) => master.id);
	const renderQuick = () => {
		if (quick.hidden) return quick.replaceChildren();
		const masters = openIds()
			.map((id) => STENCILS.find((stencil) => stencil.id === id)!)
			.flatMap((stencil) => quickIds(stencil).map((id) => findMaster(id)!.master));
		const empty = doc.createElement('p');
		empty.className = 'shapes-empty';
		empty.textContent = 'No Quick Shapes. Right-click a master to add it.';
		quick.replaceChildren(masters.length ? masterList(doc, masters) : empty);
	};
	const renderSections = () => {
		const collapsed = new Set(
			[...sections.querySelectorAll<HTMLElement>('[data-stencil][data-collapsed]')].map(
				(section) => section.dataset.stencil!,
			),
		);
		sections.replaceChildren(
			...openIds().map((id) => {
				const stencil = STENCILS.find((candidate) => candidate.id === id)!;
				const section = doc.createElement('section');
				section.dataset.stencil = id;
				section.setAttribute('aria-label', stencil.name);
				const header = doc.createElement('div');
				header.className = 'stencil-header';
				const title = doc.createElement('button');
				title.type = 'button';
				title.className = 'stencil-title';
				title.textContent = stencil.name;
				const list = masterList(doc, stencil.masters);
				list.id = `stencil-${id}-masters`;
				title.setAttribute('aria-controls', list.id);
				const open = !collapsed.has(id);
				title.setAttribute('aria-expanded', String(open));
				list.hidden = !open;
				if (!open) section.dataset.collapsed = '';
				header.append(title);
				if (id !== BASIC_STENCIL_ID) {
					const close = doc.createElement('button');
					close.type = 'button';
					close.className = 'stencil-close';
					close.dataset.closeStencil = id;
					close.setAttribute('aria-label', `Close ${stencil.name}`);
					close.title = `Close ${stencil.name}`;
					close.textContent = '×';
					header.append(close);
				}
				section.append(header, list);
				return section;
			}),
		);
	};
	const renderMenu = () =>
		moreMenu.replaceChildren(
			...STENCILS.filter((stencil) => stencil.id !== BASIC_STENCIL_ID).map((stencil) =>
				menuItem(doc, `stencil:${stencil.id}`, stencil.name, saved.open.includes(stencil.id)),
			),
			doc.createElement('office-ui-menu-separator'),
			unsupportedItem(
				menuItem(doc, 'open-stencil-file', 'Open Stencil...'),
				'Open Stencil: stencil files (.vssx, .vss) are not supported.',
			),
		);
	const render = () => {
		renderMenu();
		renderSections();
		renderQuick();
	};
	const toggleStencil = (id: string) => {
		saved.open = saved.open.includes(id)
			? saved.open.filter((candidate) => candidate !== id)
			: [...saved.open, id];
		save(doc, saved);
		render();
	};

	more.addEventListener('click', () => {
		const box = more.getBoundingClientRect();
		moreMenu.openAt?.(box.left, box.bottom);
	});
	moreMenu.addEventListener('office-command', (event) => {
		event.stopPropagation();
		const command = (event as CustomEvent<{ command: string }>).detail.command;
		if (command.startsWith('stencil:')) toggleStencil(command.slice('stencil:'.length));
	});
	quickToggle.addEventListener('click', () => {
		quick.hidden = !quick.hidden;
		quickToggle.setAttribute('aria-expanded', String(!quick.hidden));
		renderQuick();
	});
	sections.addEventListener('click', (event) => {
		const target = event.target as Element;
		const close = target.closest<HTMLElement>('[data-close-stencil]');
		if (close) return toggleStencil(close.dataset.closeStencil!);
		const title = target.closest<HTMLElement>('.stencil-title');
		if (!title) return;
		const section = title.closest<HTMLElement>('[data-stencil]')!;
		const open = title.getAttribute('aria-expanded') !== 'true';
		title.setAttribute('aria-expanded', String(open));
		section.querySelector<HTMLElement>('.masters')!.hidden = !open;
		section.toggleAttribute('data-collapsed', !open);
	});
	pane.addEventListener('contextmenu', (event) => {
		const button = (event.target as Element).closest?.<HTMLElement>('[data-master]');
		const found = button && findMaster(button.dataset.master!);
		if (!found) return;
		event.preventDefault();
		menuMaster = found.master.id;
		const added = quickIds(found.stencil).includes(menuMaster);
		masterMenu.replaceChildren(
			menuItem(doc, 'quick-shapes', added ? 'Remove from Quick Shapes' : 'Add to Quick Shapes'),
		);
		masterMenu.openAt?.(event.clientX, event.clientY);
	});
	masterMenu.addEventListener('office-command', (event) => {
		event.stopPropagation();
		const found = menuMaster && findMaster(menuMaster);
		if (!found) return;
		const ids = quickIds(found.stencil);
		saved.quick[found.stencil.id] = ids.includes(found.master.id)
			? ids.filter((id) => id !== found.master.id)
			: [...ids, found.master.id];
		save(doc, saved);
		renderQuick();
	});
	render();
	panel.append(more, quickToggle, quick, sections);
	// The caller puts the menus on the pane, so they open from the Search view too.
	return [moreMenu, masterMenu];
}
