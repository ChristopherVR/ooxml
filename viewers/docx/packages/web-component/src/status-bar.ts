import { defineStatusBar, defineZoomSlider } from 'ooxml-ui/controls';
import type { OfficeStatusBarState } from 'ooxml-ui/controls';
import { registerIcon } from 'ooxml-ui/icons';
import { normalizeEditorLocale, translate } from './localization';

export interface StatusBarHandlers {
	setViewMode(mode: 'draft' | 'print'): void;
	setZoom(percent: number): void;
	showCompatibilityNotes(): void;
}

export interface StatusBar {
	element: HTMLElement;
	setPageAndWords(page: string, words: string): void;
	setLanguage(language: string): void;
	setNoteCount(count: number): void;
	setViewMode(mode: 'draft' | 'print'): void;
	setZoom(percent: number): void;
	/** Redraws every label after the editor locale changed. */
	relocalize(): void;
}

export const ZOOM_MIN = 50;
export const ZOOM_MAX = 200;

type StatusBarElement = HTMLElement & { state: OfficeStatusBarState };
type ZoomSliderElement = HTMLElement & { value: number };

/** Word's layout switches, drawn on the shared 20px grid; product icons stay with the product. */
const LAYOUT_ICONS = {
	webLayout: 'M3 5h14v10H3z M3 8h14',
	printLayout: 'M5 3h10v14H5z M8 7h4 M8 10h4',
};

/**
 * Word-style status bar: page/words/language on the left, views and zoom on the right. The shared
 * `office-ui-status-bar` draws it from translated state and `office-ui-zoom-slider` supplies the zoom
 * controls; this module only holds the model and routes the activations to the editor.
 */
export function createStatusBar(handlers: StatusBarHandlers): StatusBar {
	defineStatusBar();
	defineZoomSlider();
	for (const [name, d] of Object.entries(LAYOUT_ICONS)) registerIcon(name, d);

	const element = document.createElement('office-ui-status-bar') as StatusBarElement;
	element.className = 'dve-status';
	const slider = document.createElement('office-ui-zoom-slider') as ZoomSliderElement;
	slider.slot = 'end';
	slider.setAttribute('min', String(ZOOM_MIN));
	slider.setAttribute('max', String(ZOOM_MAX));
	slider.className = 'dve-zoom';
	element.append(slider);

	const model = { page: '', words: '', language: '', notes: 0, mode: 'draft', zoom: 100 };
	const locale = () => normalizeEditorLocale(element.dataset.editorLocale);

	const render = () => {
		const t = locale();
		const noteKey = model.notes === 1 ? 'status.note' : 'status.notes';
		const layout = (mode: 'draft' | 'print', key: 'Web Layout' | 'Print Layout') => ({
			id: mode,
			icon: mode === 'draft' ? 'webLayout' : 'printLayout',
			label: translate(t, key),
			pressed: model.mode === mode,
		});
		element.state = {
			items: [
				{
					id: 'page',
					text: model.page,
					title: translate(t, 'nav.approximate'),
					live: true,
				},
				{ id: 'words', text: model.words },
				...(model.language ? [{ id: 'language', text: model.language, narrowHide: true }] : []),
			],
			toggles: [
				{
					id: 'notes',
					icon: 'warning',
					label: translate(t, noteKey).replace('{count}', String(model.notes)),
					text: translate(t, noteKey).replace('{count}', String(model.notes)),
					hidden: model.notes === 0,
				},
			],
			views: [layout('draft', 'Web Layout'), layout('print', 'Print Layout')],
		};
		slider.setAttribute('out-label', translate(t, 'Zoom out'));
		slider.setAttribute('in-label', translate(t, 'Zoom in'));
		slider.setAttribute('slider-label', translate(t, 'Zoom level'));
		slider.value = model.zoom;
	};

	element.addEventListener('office-status-activate', (event) => {
		const { id } = (event as CustomEvent<{ id: string }>).detail;
		if (id === 'notes') handlers.showCompatibilityNotes();
		else if (id === 'draft' || id === 'print') handlers.setViewMode(id);
	});
	slider.addEventListener('input', () => handlers.setZoom(slider.value));

	render();
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
