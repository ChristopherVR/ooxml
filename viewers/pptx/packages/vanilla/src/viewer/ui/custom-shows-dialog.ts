import { createCustomShowsController } from 'ooxml-ui/pptx';
import type { CustomShowsSlideRow } from 'ooxml-ui/pptx';
import type { PptxCustomShow, PptxSlide } from 'pptx-viewer-core';

import type { Translator } from '../i18n';
import { createEl } from '../render';
import { appendDialogActions } from './dialog-footer';
import { createInspectorSelect } from './inspector/controls-extra';
import { createParityDialogShell } from './parity-dialog-shell';
import type { RegionSet } from './region-set';

export interface CustomShowsDialogOptions {
	shows: readonly PptxCustomShow[];
	slides: readonly PptxSlide[];
	/** The show playback is currently restricted to, or null for the whole deck. */
	activeShowId: string | null;
	/** Commit the edited shows; `slides` is set when jumps into removed shows were cleared. */
	onSave(shows: PptxCustomShow[], slides: PptxSlide[] | undefined): void;
	/** Restrict (or, with null, stop restricting) the slide show to one show. */
	onSetActive(id: string | null): void;
	/** Start the show, already restricted to `show`. */
	onRun(show: PptxCustomShow): void;
	/** The viewer's dialog regions: the dialog closes with the viewer. */
	regions?: RegionSet;
}

function button(
	doc: Document,
	label: string,
	run: () => void,
	disabled = false,
): HTMLButtonElement {
	const el = createEl(doc, 'button');
	el.type = 'button';
	el.textContent = label;
	el.disabled = disabled;
	el.addEventListener('click', run);
	return el;
}

/**
 * The Custom Shows dialog (Slide Show > Custom Slide Show): New, Copy and Remove for the
 * shows, and for the selected one PowerPoint's two lists (slides in the presentation, slides in
 * the custom show) with Add, Remove and Move Up/Down. All of its logic is the shared
 * `createCustomShowsController`; this file only draws it.
 */
export function openCustomShowsDialog(
	doc: Document,
	t: Translator,
	options: CustomShowsDialogOptions,
): void {
	const { slides, onSave, onSetActive, onRun } = options;
	const shell = createParityDialogShell(
		doc,
		t,
		t('pptx.customShows.title'),
		options.regions && { regions: options.regions },
	);
	const controller = createCustomShowsController(options.shows, options.activeShowId);
	let activeShowId = options.activeShowId;

	const picker = createEl(doc, 'label', 'pptxv-custom-shows-active');
	const pickerLabel = createEl(doc, 'span');
	pickerLabel.textContent = t('pptx.customShows.selectCustomShow');
	const select = createInspectorSelect(doc);
	select.setAttribute('aria-label', t('pptx.customShows.selectCustomShow'));
	select.addEventListener('change', () => {
		activeShowId = select.value || null;
		onSetActive(activeShowId);
	});
	picker.append(pickerLabel, select);

	const layout = createEl(doc, 'div', 'pptxv-custom-shows');
	const showsColumn = createEl(doc, 'div', 'pptxv-custom-shows-list');
	const editor = createEl(doc, 'div', 'pptxv-custom-shows-editor');
	layout.append(showsColumn, editor);
	shell.body.append(picker, layout);

	const nameLabel = createEl(doc, 'label');
	const nameText = createEl(doc, 'span');
	nameText.textContent = t('pptx.customShows.renameNameLabel');
	const nameInput = createEl(doc, 'input');
	nameInput.setAttribute('aria-label', t('pptx.customShows.renameNameLabel'));
	nameInput.addEventListener('change', () => controller.rename(nameInput.value));
	nameLabel.append(nameText, nameInput);
	const lists = createEl(doc, 'div', 'pptxv-custom-shows-lists');
	editor.append(nameLabel, lists);

	const rows = (
		heading: string,
		list: readonly CustomShowsSlideRow[],
		toggle: (row: CustomShowsSlideRow) => void,
		dbl?: (row: CustomShowsSlideRow) => void,
	): HTMLElement => {
		const column = createEl(doc, 'div');
		const title = createEl(doc, 'strong');
		title.textContent = heading;
		const ul = createEl(doc, 'ul');
		ul.setAttribute('role', 'listbox');
		ul.setAttribute('aria-label', heading);
		ul.setAttribute('aria-multiselectable', 'true');
		for (const row of list) {
			const li = createEl(doc, 'li');
			li.setAttribute('role', 'option');
			li.setAttribute('aria-selected', String(row.selected));
			li.tabIndex = 0;
			li.textContent = `${row.number}. ${row.title}`.trim();
			li.addEventListener('click', () => toggle(row));
			li.addEventListener('keydown', (event) => {
				if (event.key === ' ' || event.key === 'Enter') {
					event.preventDefault();
					toggle(row);
				}
			});
			if (dbl) {
				li.addEventListener('dblclick', () => dbl(row));
			}
			ul.appendChild(li);
		}
		column.append(title, ul);
		return column;
	};

	const render = (): void => {
		const view = controller.view(slides);
		showsColumn.replaceChildren();
		const ul = createEl(doc, 'ul');
		ul.setAttribute('role', 'listbox');
		ul.setAttribute('aria-label', t('pptx.customShows.savedShows'));
		for (const show of view.shows) {
			const li = createEl(doc, 'li');
			li.setAttribute('role', 'option');
			li.setAttribute('aria-selected', String(show.selected));
			li.tabIndex = 0;
			li.textContent = show.name;
			li.addEventListener('click', () => controller.select(show.id));
			ul.appendChild(li);
		}
		if (view.shows.length === 0) {
			const empty = createEl(doc, 'p');
			empty.textContent = t('pptx.customShows.empty');
			showsColumn.appendChild(empty);
		}
		showsColumn.append(
			ul,
			button(doc, t('pptx.customShows.createNew'), () =>
				controller.create(t('pptx.customShows.newDefaultName')),
			),
			button(doc, t('pptx.customShows.copy'), () => controller.copySelected(), !view.canEdit),
			button(doc, t('pptx.customShows.remove'), () => controller.removeSelected(), !view.canEdit),
		);

		editor.hidden = !view.selected;
		if (view.selected && doc.activeElement !== nameInput) {
			nameInput.value = view.selected.name;
		}
		const middle = createEl(doc, 'div', 'pptxv-custom-shows-actions');
		middle.append(
			button(doc, t('pptx.customShows.add'), () => controller.addSelected(slides), !view.canAdd),
			button(
				doc,
				t('pptx.customShows.remove'),
				() => controller.removeSelectedSlides(),
				!view.canRemoveSlides,
			),
		);
		const order = createEl(doc, 'div', 'pptxv-custom-shows-actions');
		order.append(
			button(doc, t('pptx.customShows.moveUp'), () => controller.moveUp(), !view.canMoveUp),
			button(doc, t('pptx.customShows.moveDown'), () => controller.moveDown(), !view.canMoveDown),
		);
		lists.replaceChildren(
			rows(
				t('pptx.customShows.slidesInPresentation'),
				view.presentationSlides,
				(row) => controller.toggleLeft(row.rId),
				(row) => controller.addSlide(row.rId),
			),
			middle,
			rows(t('pptx.customShows.slidesInShow'), view.showSlides, (row) =>
				controller.toggleRight(row.position ?? 0),
			),
			order,
		);

		select.replaceChildren();
		const all = doc.createElement('option');
		all.value = '';
		all.textContent = t('pptx.customShows.allSlides');
		select.appendChild(all);
		for (const show of view.shows) {
			const option = doc.createElement('option');
			option.value = show.id;
			option.textContent = show.name;
			select.appendChild(option);
		}
		// A show can be removed in this same dialog, so the selection is re-derived.
		select.value = view.shows.some(({ id }) => id === activeShowId) ? (activeShowId ?? '') : '';
	};
	controller.subscribe(render);

	const commit = (): void => {
		const result = controller.result(slides);
		onSave(result.shows, result.slides);
		// A show removed in this session must not leave the slide show pinned to an id
		// nothing resolves: that would silently present the whole deck.
		onSetActive(result.shows.some(({ id }) => id === activeShowId) ? activeShowId : null);
	};

	appendDialogActions(doc, shell.footer, [
		{
			id: 'show',
			label: t('pptx.customShows.showButton'),
			align: 'start',
			run: () => {
				const id = controller.getSnapshot().selectedId;
				const show = controller.getSnapshot().shows.find((s) => s.id === id);
				if (!show) {
					return;
				}
				commit();
				// Running a show IS selecting it: the show that starts must be the one whose
				// membership the navigation rule then honours.
				onSetActive(show.id);
				onRun(show);
				shell.close();
			},
		},
		{ id: 'cancel', label: t('pptx.common.cancel'), run: shell.close },
		{
			id: 'ok',
			label: t('pptx.common.ok'),
			variant: 'primary',
			run: () => {
				commit();
				shell.close();
			},
		},
	]);
	render();
}
