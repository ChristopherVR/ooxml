import { DOCUMENT_VIEWS, isDocumentView, type DocumentView } from './document-views';
import { defineStatusBar, defineZoomSlider } from '../controls';
import type { OfficeStatusBarState } from '../controls';
import { registerIcon } from '../icons';
import {
	compactPageText,
	statusLayout,
	watchStatusWidth,
	type StatusLayout,
} from '../chrome/status-layout';
import { AUTO_FIT_MIN } from './auto-fit';
import { normalizeEditorLocale, translate } from './localization';

export interface StatusBarHandlers {
	setViewMode(mode: DocumentView): void;
	setZoom(percent: number): void;
	showCompatibilityNotes(): void;
}

export interface StatusBar {
	element: HTMLElement;
	setPageAndWords(page: string, words: string): void;
	setLanguage(language: string): void;
	setNoteCount(count: number): void;
	/** False while the File tab is hidden: the notes button opens File > Info, so it goes too. */
	setNotesAvailable(available: boolean): void;
	setViewMode(mode: DocumentView): void;
	setZoom(percent: number): void;
	/** Redraws every label after the editor locale changed. */
	relocalize(): void;
}

export const ZOOM_MIN = AUTO_FIT_MIN;
export const ZOOM_MAX = 200;

type StatusBarElement = HTMLElement & { state: OfficeStatusBarState };
type ZoomSliderElement = HTMLElement & { value: number };

/** Word's layout switches, drawn on the shared 20px grid; product icons stay with the product. */
const LAYOUT_ICONS = {
	draft: 'M4 5h12 M4 8h12 M4 11h9 M4 14h10',
	read: 'M10 6C8 4 5 4 2 5v10c3-1 6-1 8 1 2-2 5-2 8-1V5c-3-1-6-1-8 1z M10 6v10',
	outline: 'M3 5h2 M7 5h10 M5 9h2 M9 9h8 M5 13h2 M9 13h8',
	focus: 'M3 8V3h5 M12 3h5v5 M17 12v5h-5 M8 17H3v-5',
	web: 'M3 5h14v10H3z M3 8h14',
	print: 'M5 3h10v14H5z M8 7h4 M8 10h4',
};

/**
 * Word-style status bar: page/words/language on the left, views and zoom on the right. The shared
 * `office-ui-status-bar` draws it from translated state and `office-ui-zoom-slider` supplies the zoom
 * controls; this module only holds the model and routes the activations to the editor.
 */
export function createStatusBar(handlers: StatusBarHandlers): StatusBar {
	defineStatusBar();
	defineZoomSlider();
	for (const [name, d] of Object.entries(LAYOUT_ICONS)) registerIcon(`word-view-${name}`, d);

	const element = document.createElement('office-ui-status-bar') as StatusBarElement;
	element.className = 'dve-status';
	const slider = document.createElement('office-ui-zoom-slider') as ZoomSliderElement;
	slider.slot = 'end';
	slider.setAttribute('min', String(ZOOM_MIN));
	slider.setAttribute('max', String(ZOOM_MAX));
	slider.className = 'dve-zoom';
	element.append(slider);

	const model = {
		page: '',
		words: '',
		language: '',
		notes: 0,
		notesAvailable: true,
		mode: 'draft' as DocumentView,
		zoom: 100,
	};
	let layout: StatusLayout = statusLayout(0);
	const locale = () => normalizeEditorLocale(element.dataset.editorLocale);

	const render = () => {
		const t = locale();
		const noteKey = model.notes === 1 ? 'status.note' : 'status.notes';
		const view = ([mode, key]: (typeof DOCUMENT_VIEWS)[number]) => ({
			id: mode,
			icon: `word-view-${mode}`,
			label: translate(t, key),
			pressed: model.mode === mode,
		});
		element.state = {
			items: [
				{
					id: 'page',
					text: layout.compactPage ? compactPageText(model.page) : model.page,
					title: layout.compactPage
						? `${model.page}. ${translate(t, 'nav.approximate')}`
						: translate(t, 'nav.approximate'),
					live: true,
				},
				...(layout.words ? [{ id: 'words', text: model.words }] : []),
				...(model.language && layout.language
					? [{ id: 'language', text: model.language, narrowHide: true }]
					: []),
			],
			toggles: [
				{
					id: 'notes',
					icon: 'warning',
					label: translate(t, noteKey).replace('{count}', String(model.notes)),
					text: translate(t, noteKey).replace('{count}', String(model.notes)),
					hidden: model.notes === 0 || !model.notesAvailable,
				},
			],
			views: layout.views ? DOCUMENT_VIEWS.map(view) : false,
		};
		slider.setAttribute('out-label', translate(t, 'Zoom out'));
		slider.setAttribute('in-label', translate(t, 'Zoom in'));
		slider.setAttribute('slider-label', translate(t, 'Zoom level'));
		slider.value = model.zoom;
	};

	element.addEventListener('office-status-activate', (event) => {
		const { id } = (event as CustomEvent<{ id: string }>).detail;
		if (id === 'notes') handlers.showCompatibilityNotes();
		else if (isDocumentView(id)) handlers.setViewMode(id);
	});
	slider.addEventListener('input', () => handlers.setZoom(slider.value));

	render();
	watchStatusWidth(element, (width) => {
		layout = statusLayout(width);
		render();
	});
	return {
		element,
		setPageAndWords(pageText, wordText) {
			model.page = pageText;
			model.words = wordText;
			render();
		},
		setLanguage(value) {
			model.language = value;
			render();
		},
		setNoteCount(count) {
			model.notes = count;
			render();
		},
		setNotesAvailable(available) {
			if (model.notesAvailable === available) return;
			model.notesAvailable = available;
			render();
		},
		setViewMode(mode) {
			model.mode = mode;
			render();
		},
		setZoom(value) {
			model.zoom = value;
			render();
		},
		relocalize: render,
	};
}
