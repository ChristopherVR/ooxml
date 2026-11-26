import { createStencilSection } from './shapes-section';
import { rememberMasterFocus, rememberSectionFocus } from './shapes-focus';
import { BASIC_STENCIL_ID, STENCILS, findMaster, type Stencil } from './stencil-catalog';
import {
	DOCUMENT_STENCIL_ID,
	DOCUMENT_STENCIL_NAME,
	EMPTY_SHAPES_DOCUMENT,
	unavailableStencil,
	type ShapesDocument,
} from './shapes-document';
import { loadShapes as load, saveShapes as save } from './shapes-storage';
import { MY_SHAPES_ID, stencilFiles } from './stencil-files';
import { wireStencilMenus } from './shapes-stencil-menus';
import type { StencilMenu } from './stencil-menus';
import { masterList, menuItem } from './shapes-master-list';
import { stencilFilePicker } from './shapes-files-ui';
import { translateStencilText } from './stencil-locale';
import type { ViewerController } from './controller';
export { SHAPES_STORAGE_KEY, currentQuickShapes } from './shapes-storage';

/** The Stencils view of one Shapes window. */
export interface StencilsView {
	/** The More Shapes and master menus; the caller puts them on the pane. */
	menus: HTMLElement[];
	refreshLocale(): void;
	/**
	 * Show a drawing's document stencil and the stencils it docks; a no-op for the same content.
	 * `opened` is false after an edit of the drawing that is showing: folded and unfolded stencils
	 * stay as the user has them, and a Document Stencil that appears then starts folded.
	 */
	setDocument(next: ShapesDocument, opened?: boolean): void;
}

/** Visio shows a stencil's first masters as its Quick Shapes until the user changes them. */
const DEFAULT_QUICK = 4;

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
	controller: () => ViewerController | undefined,
): StencilsView {
	const saved = load(doc);
	const files = stencilFiles(pane);
	let drawing: ShapesDocument = EMPTY_SHAPES_DOCUMENT;
	/** Docked stencils the user has not closed in this session. */
	let docked: string[] = [];
	const collapsed = new Set<string>();
	const row = (id: string) => {
		const button = doc.createElement('button');
		button.type = 'button';
		button.id = id;
		button.className = 'shapes-row';
		return button;
	};
	const more = row('shapes-more');
	more.setAttribute('aria-haspopup', 'menu');
	const quickToggle = row('shapes-quick-toggle');
	quickToggle.setAttribute('aria-expanded', 'false');
	quickToggle.setAttribute('aria-controls', 'shapes-quick');
	const quick = doc.createElement('div');
	quick.id = 'shapes-quick';
	quick.hidden = true;
	const sections = doc.createElement('div');
	sections.id = 'shapes-sections';

	const t = (text: string) => translateStencilText(pane, text);
	const picker = stencilFilePicker(
		doc,
		files,
		(stencil) => {
			collapsed.delete(stencil.id);
			render();
		},
		t,
	);
	const moreMenu = doc.createElement('office-ui-context-menu') as StencilMenu;
	moreMenu.dataset.shapesMenu = 'more';
	const masterMenu = doc.createElement('office-ui-context-menu') as StencilMenu;
	masterMenu.dataset.shapesMenu = 'master';
	const localize = () => {
		more.textContent = t('More Shapes');
		more.title = t('More Shapes: open another stencil.');
		quickToggle.textContent = t('Quick Shapes');
		quickToggle.title = t('Quick Shapes: favourite masters of every open stencil.');
		moreMenu.setAttribute('label', more.textContent);
		masterMenu.setAttribute('label', t('Master'));
		picker.input.setAttribute('aria-label', t('Open Stencil'));
	};

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
	const showError = (message: string) => {
		picker.status.textContent = message;
		picker.status.hidden = !message;
	};
	const authoring = wireStencilMenus(pane, files, masterMenu, {
		controller,
		quickIds,
		changeQuick: (stencil, id) => {
			const ids = quickIds(stencil);
			saved.quick[stencil.id] = ids.includes(id)
				? ids.filter((candidate) => candidate !== id)
				: [...ids, id];
			save(doc, saved);
			renderQuick();
		},
		error: showError,
		update(id) {
			collapsed.delete(id);
			render();
		},
	});

	const renderQuick = () => {
		if (quick.hidden) return quick.replaceChildren();
		const restoreFocus = rememberMasterFocus(quick, quickToggle);
		const masters = [
			// The drawing's own masters lead, as its Document Stencil leads the stencils.
			...drawing.masters.filter((master) => !master.unsupported).slice(0, DEFAULT_QUICK),
			...openIds()
				.map((id) => STENCILS.find((stencil) => stencil.id === id)!)
				.flatMap((stencil) => quickIds(stencil).map((id) => findMaster(id)!.master)),
		];
		const empty = doc.createElement('p');
		empty.className = 'shapes-empty';
		empty.textContent = t('No Quick Shapes. Right-click a master to add it.');
		quick.replaceChildren(masters.length ? masterList(doc, masters, t) : empty);
		restoreFocus?.();
	};
	const renderSections = () => {
		const restoreFocus = rememberSectionFocus(sections, more);
		sections.replaceChildren(
			...sectionIds().map((id) => {
				const stencil = sectionStencil(id);
				const file = files.get(id);
				return createStencilSection(doc, {
					stencil,
					name: file && id !== MY_SHAPES_ID ? stencil.name : t(stencil.name),
					open: !collapsed.has(id),
					editable: file?.editable ?? false,
					omitted: id === DOCUMENT_STENCIL_ID ? drawing.omitted : 0,
					t,
				});
			}),
			// Stencil files this viewer cannot open are still named, as Visio lists them.
			...drawing.unavailable.map((file) => unavailableStencil(doc, file, t)),
		);
		restoreFocus?.();
	};
	const renderMenu = () =>
		moreMenu.replaceChildren(
			...STENCILS.filter((stencil) => stencil.id !== BASIC_STENCIL_ID).map((stencil) =>
				menuItem(doc, `stencil:${stencil.id}`, t(stencil.name), openIds().includes(stencil.id)),
			),
			doc.createElement('office-ui-menu-separator'),
			menuItem(doc, 'new-stencil', t('New Stencil')),
			Object.assign(menuItem(doc, 'open-stencil-file', t('Open Stencil...')), {
				title: t('Open Stencil: a .vssx stencil file. Legacy .vss files are not supported.'),
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
		moreMenu.openAt?.(box.left, box.bottom, more);
	});
	moreMenu.addEventListener('office-command', (event) => {
		event.stopPropagation();
		const command = (event as CustomEvent<{ command: string }>).detail.command;
		if (command.startsWith('stencil:')) toggleStencil(command.slice('stencil:'.length));
		else if (command === 'open-stencil-file') picker.open();
		else if (command === 'new-stencil') {
			moreMenu.close?.();
			authoring.create();
		}
	});
	quickToggle.addEventListener('click', () => {
		quick.hidden = !quick.hidden;
		quickToggle.setAttribute('aria-expanded', String(!quick.hidden));
		renderQuick();
	});
	sections.addEventListener('click', (event) => {
		const target = event.target as Element;
		const action = target.closest<HTMLElement>('[data-stencil-actions]');
		const id = action?.dataset.stencilActions;
		if (action && id) return authoring.openSection(action, id);
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

	localize();
	render();
	panel.append(more, quickToggle, quick, picker.status, sections, picker.input);
	return {
		// The caller puts the menus on the pane, so they open from the Search view too.
		menus: [moreMenu, masterMenu],
		refreshLocale() {
			localize();
			render();
		},
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
