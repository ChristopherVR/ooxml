import { interpolate } from 'ooxml-core/i18n';
import { BASIC_STENCIL_ID, STENCILS, findMaster, type Stencil } from './stencil-catalog';
import {
	DOCUMENT_MASTER_LIMIT,
	DOCUMENT_STENCIL_ID,
	DOCUMENT_STENCIL_NAME,
	EMPTY_SHAPES_DOCUMENT,
	unavailableStencil,
	type ShapesDocument,
} from './shapes-document';
import { loadShapes as load, saveShapes as save } from './shapes-storage';
import { MY_SHAPES_ID, stencilFiles } from './stencil-files';
import { bindStencilMenus } from './stencil-menus';
import { masterList, menuItem } from './shapes-master-list';
export { masterButton, masterList } from './shapes-master-list';
import { stencilFilePicker } from './shapes-files-ui';
import { ribbonLocale } from './ribbon-localize';
import { translateRibbonText } from './ribbon-locale';
export { SHAPES_STORAGE_KEY, currentQuickShapes } from './shapes-storage';

/** The Stencils view of one Shapes window. */
export interface StencilsView {
	/** The More Shapes and master menus; the caller puts them on the pane. */
	menus: HTMLElement[];
	/**
	 * Show a drawing's document stencil and the stencils it docks; a no-op for the same content.
	 * `opened` is false after an edit of the drawing that is showing: folded and unfolded stencils
	 * stay as the user has them, and a Document Stencil that appears then starts folded.
	 */
	setDocument(next: ShapesDocument, opened?: boolean): void;
}

/** Visio shows a stencil's first masters as its Quick Shapes until the user changes them. */
const DEFAULT_QUICK = 4;
/** Undefined until the shared menu element is registered. */
type Menu = HTMLElement & { openAt?(x: number, y: number): void };

/**
 * The Stencils view below the view tabs: More Shapes (a menu of the built-in stencils), Quick
 * Shapes (favourites across the open stencils) and one collapsible section per open stencil.
 * A drawing's own masters come first as its Document Stencil, then the stencils the drawing docks,
 * then Basic Shapes and the stencils the user opened. The first of the drawing's stencils is the
 * one showing, as in Visio; the others start folded. Opened stencils and Quick Shapes persist per
 * viewer in localStorage; what a drawing docks is not persisted.
 */
export function buildStencilsView(
	doc: Document,
	panel: HTMLElement,
	pane: HTMLElement,
): StencilsView {
	const saved = load(doc);
	const files = stencilFiles(pane);
	let drawing: ShapesDocument = EMPTY_SHAPES_DOCUMENT;
	/** Docked stencils the user has not closed in this session. */
	let docked: string[] = [];
	const collapsed = new Set<string>();
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

	const t = (text: string) => translateRibbonText(ribbonLocale(pane), text);
	const picker = stencilFilePicker(
		doc,
		files,
		(stencil) => {
			collapsed.delete(stencil.id);
			render();
		},
		t,
	);
	const moreMenu = doc.createElement('office-ui-context-menu') as Menu;
	moreMenu.dataset.shapesMenu = 'more';
	moreMenu.setAttribute('label', 'More Shapes');
	const masterMenu = doc.createElement('office-ui-context-menu') as Menu;
	masterMenu.dataset.shapesMenu = 'master';
	masterMenu.setAttribute('label', 'Master');

	const openIds = () => [...new Set([...docked, BASIC_STENCIL_ID, ...saved.open])];
	const fileSections = () => files.opened.map((stencil) => stencil.id);
	const sectionIds = () => [
		...(drawing.masters.length ? [DOCUMENT_STENCIL_ID] : []),
		...openIds(),
		...fileSections(),
	];
	const sectionStencil = (id: string): Stencil =>
		id === DOCUMENT_STENCIL_ID
			? { id, name: DOCUMENT_STENCIL_NAME, masters: drawing.masters }
			: (files.get(id) ?? STENCILS.find((candidate) => candidate.id === id)!);
	const quickIds = (stencil: Stencil) =>
		saved.quick[stencil.id] ?? stencil.masters.slice(0, DEFAULT_QUICK).map((master) => master.id);
	const authoring = bindStencilMenus(pane, files, masterMenu, {
		quickIds,
		changeQuick: (stencil, id) => {
			const ids = quickIds(stencil);
			saved.quick[stencil.id] = ids.includes(id)
				? ids.filter((candidate) => candidate !== id)
				: [...ids, id];
			save(doc, saved);
			renderQuick();
		},
		error(message) {
			picker.status.textContent = message;
			picker.status.hidden = !message;
		},
		update(id) {
			collapsed.delete(id);
			render();
			Array.from(sections.querySelectorAll<HTMLButtonElement>('[data-stencil-actions]'))
				.find((button) => button.dataset.stencilActions === id)
				?.focus();
		},
	});

	const renderQuick = () => {
		if (quick.hidden) return quick.replaceChildren();
		const masters = [
			// The drawing's own masters lead, as its Document Stencil leads the stencils.
			...drawing.masters.filter((master) => !master.unsupported).slice(0, DEFAULT_QUICK),
			...openIds()
				.map((id) => STENCILS.find((stencil) => stencil.id === id)!)
				.flatMap((stencil) => quickIds(stencil).map((id) => findMaster(id)!.master)),
		];
		const empty = doc.createElement('p');
		empty.className = 'shapes-empty';
		empty.textContent = 'No Quick Shapes. Right-click a master to add it.';
		quick.replaceChildren(masters.length ? masterList(doc, masters) : empty);
	};
	const renderSections = () => {
		sections.replaceChildren(
			...sectionIds().map((id) => {
				const stencil = sectionStencil(id);
				const section = doc.createElement('section');
				section.dataset.stencil = id;
				const name = id === MY_SHAPES_ID ? t(stencil.name) : stencil.name;
				section.setAttribute('aria-label', name);
				const header = doc.createElement('div');
				header.className = 'stencil-header';
				const title = doc.createElement('button');
				title.type = 'button';
				title.className = 'stencil-title';
				title.textContent = name;
				const list = masterList(doc, stencil.masters);
				list.id = `stencil-${id}-masters`;
				title.setAttribute('aria-controls', list.id);
				const open = !collapsed.has(id);
				title.setAttribute('aria-expanded', String(open));
				list.hidden = !open;
				if (!open) section.dataset.collapsed = '';
				header.append(title);
				if (files.get(id)?.editable) {
					const actions = doc.createElement('button');
					actions.className = 'stencil-close';
					actions.type = 'button';
					actions.textContent = '⋯';
					actions.setAttribute('aria-haspopup', 'menu');
					actions.setAttribute('aria-label', t('Stencil actions'));
					actions.title = t('Stencil actions');
					actions.dataset.stencilActions = id;
					header.append(actions);
				}
				if (id !== BASIC_STENCIL_ID && id !== DOCUMENT_STENCIL_ID) {
					const close = doc.createElement('button');
					close.type = 'button';
					close.className = 'stencil-close';
					close.dataset.closeStencil = id;
					const closeLabel = interpolate(t('Close {name}'), { name });
					close.setAttribute('aria-label', closeLabel);
					close.title = closeLabel;
					close.textContent = '×';
					header.append(close);
				}
				section.append(header, list);
				if (id === DOCUMENT_STENCIL_ID && drawing.omitted) {
					const note = doc.createElement('p');
					note.className = 'shapes-empty';
					note.hidden = !open;
					note.textContent = `Showing the first ${DOCUMENT_MASTER_LIMIT} masters; ${drawing.omitted} more are not listed.`;
					section.append(note);
				}
				return section;
			}),
			// Stencil files this viewer cannot open are still named, as Visio lists them.
			...drawing.unavailable.map((file) => unavailableStencil(doc, file)),
		);
	};
	const renderMenu = () =>
		moreMenu.replaceChildren(
			...STENCILS.filter((stencil) => stencil.id !== BASIC_STENCIL_ID).map((stencil) =>
				menuItem(doc, `stencil:${stencil.id}`, stencil.name, openIds().includes(stencil.id)),
			),
			doc.createElement('office-ui-menu-separator'),
			menuItem(doc, 'new-stencil', t('New Stencil')),
			Object.assign(menuItem(doc, 'open-stencil-file', 'Open Stencil...'), {
				title: 'Open Stencil: a .vssx stencil file. Legacy .vss files are not supported.',
			}),
		);
	const render = () => {
		renderMenu();
		renderSections();
		renderQuick();
	};
	const toggleStencil = (id: string) => {
		if (files.opened.some((stencil) => stencil.id === id)) {
			files.close(id);
			authoring.close(id);
			collapsed.delete(id);
			return render();
		}
		if (openIds().includes(id)) {
			docked = docked.filter((candidate) => candidate !== id);
			saved.open = saved.open.filter((candidate) => candidate !== id);
		} else saved.open = [...saved.open, id];
		collapsed.delete(id);
		save(doc, saved);
		render();
	};

	more.addEventListener('click', () => {
		const box = more.getBoundingClientRect();
		renderMenu();
		moreMenu.openAt?.(box.left, box.bottom);
	});
	moreMenu.addEventListener('office-command', (event) => {
		event.stopPropagation();
		const command = (event as CustomEvent<{ command: string }>).detail.command;
		if (command.startsWith('stencil:')) toggleStencil(command.slice('stencil:'.length));
		else if (command === 'open-stencil-file') picker.open();
		else if (command === 'new-stencil') authoring.create();
	});
	quickToggle.addEventListener('click', () => {
		quick.hidden = !quick.hidden;
		quickToggle.setAttribute('aria-expanded', String(!quick.hidden));
		renderQuick();
	});
	sections.addEventListener('click', (event) => {
		const target = event.target as Element;
		const action = target.closest<HTMLElement>('[data-stencil-actions]');
		if (action?.dataset.stencilActions) {
			authoring.openSection(action, action.dataset.stencilActions);
			return;
		}
		const close = target.closest<HTMLElement>('[data-close-stencil]');
		if (close) return toggleStencil(close.dataset.closeStencil!);
		const title = target.closest<HTMLElement>('.stencil-title');
		if (!title) return;
		const section = title.closest<HTMLElement>('[data-stencil]')!;
		const open = title.getAttribute('aria-expanded') !== 'true';
		title.setAttribute('aria-expanded', String(open));
		section.querySelector<HTMLElement>('.masters')!.hidden = !open;
		const note = section.querySelector<HTMLElement>('.shapes-empty');
		if (note) note.hidden = !open;
		section.toggleAttribute('data-collapsed', !open);
		if (open) collapsed.delete(section.dataset.stencil!);
		else collapsed.add(section.dataset.stencil!);
	});

	render();
	panel.append(more, quickToggle, quick, picker.status, sections, picker.input);
	return {
		// The caller puts the menus on the pane, so they open from the Search view too.
		menus: [moreMenu, masterMenu, authoring.menu],
		setDocument(next, opened = true) {
			// An opened drawing always starts from its own stencils, even when it lists what the
			// last one did (two drawings without masters or docked stencils).
			if (next.key === drawing.key && !opened) return;
			if (!opened) {
				// The first shape dropped in a new drawing gives it a Document Stencil: it is listed,
				// folded, and the stencil the user is working from stays open.
				const listed = new Set(sectionIds());
				drawing = next;
				for (const id of sectionIds()) if (!listed.has(id)) collapsed.add(id);
				return render();
			}
			drawing = next;
			docked = next.docked.filter((id) => STENCILS.some((stencil) => stencil.id === id));
			// The stencil the drawing docks is the one showing, as in Visio, where the Document Stencil
			// stays out of the way; a drawing that docks none of ours shows its own masters.
			const showing = docked[0] ?? (next.masters.length ? DOCUMENT_STENCIL_ID : undefined);
			collapsed.clear();
			if (showing) for (const id of sectionIds()) if (id !== showing) collapsed.add(id);
			render();
		},
	};
}
