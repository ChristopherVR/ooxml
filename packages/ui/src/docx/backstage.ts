import { defineBackstage } from '../controls';
import type { OfficeBackstageItem } from '../controls';
import type { FileCommand } from './file-commands';
import { normalizeEditorLocale, translate, translateUiText } from './localization';
import { renderHome } from './backstage-home';
import { backstageFocus } from './backstage-focus';
import {
	renderExport,
	renderInfo,
	renderNew,
	renderOpen,
	renderOptions,
	renderCustomize,
	renderPrint,
	renderSaveAs,
	type BackstageHandlers,
	type PageContext,
} from './backstage-pages';

export type { BackstageHandlers } from './backstage-pages';
export type BackstagePage =
	| 'home'
	| 'info'
	| 'new'
	| 'open'
	| 'saveAs'
	| 'print'
	| 'export'
	| 'options'
	| 'customize';

export interface Backstage {
	element: HTMLElement;
	open(page?: BackstagePage): void;
	close(): void;
	readonly isOpen: boolean;
	/** Retranslates the navigation after the editor locale changed, and redraws the open page. */
	relocalize(): void;
}

type BackstageElement = HTMLElement & {
	items: readonly OfficeBackstageItem[];
	open: boolean;
	show(page?: string): void;
};

/** Navigation order. Save sits among the file actions and has no page of its own. */
const NAV: Array<{ id: BackstagePage | 'save'; label: string; footer?: boolean }> = [
	{ id: 'home', label: 'Home' },
	{ id: 'info', label: 'Info' },
	{ id: 'new', label: 'New' },
	{ id: 'open', label: 'Open' },
	{ id: 'save', label: 'Save' },
	{ id: 'saveAs', label: 'Save As' },
	{ id: 'print', label: 'Print' },
	{ id: 'export', label: 'Export' },
	{ id: 'options', label: 'Options', footer: true },
	{ id: 'customize', label: 'Customize Ribbon', footer: true },
];

const RENDERERS: Record<BackstagePage, (context: PageContext) => void> = {
	home: renderHome,
	info: renderInfo,
	new: renderNew,
	open: renderOpen,
	saveAs: renderSaveAs,
	print: renderPrint,
	export: renderExport,
	options: renderOptions,
	customize: renderCustomize,
};

/**
 * Word's File view. Left: Back, Info, New, Open, Save and its siblings (Save As, Print, Export)
 * and Options; right: the selected page. Info shows properties and compatibility notes, Save As
 * takes a file name, Export offers PDF (through printing), DOCX and plain text, and Options sets
 * the editor's language, theme and review author.
 */
export function createBackstage(handlers: BackstageHandlers): Backstage {
	defineBackstage();
	const element = document.createElement('office-ui-backstage') as BackstageElement;
	element.className = 'dve-backstage';
	const focus = backstageFocus(element);
	const t = (text: string) => translateUiText(element, text);
	const pages = new Map<BackstagePage, HTMLElement>();
	for (const { id } of NAV) {
		if (id === 'save') continue;
		const page = document.createElement('div');
		page.className = 'dve-backstage-content';
		page.dataset.backstagePage = id;
		page.hidden = true;
		pages.set(id, page);
		element.append(page);
	}

	let current: BackstagePage = 'info';
	// A changed language or theme re-renders the open page so its text follows.
	const pageHandlers: BackstageHandlers = {
		...handlers,
		setOption(key, value) {
			const content = pages.get(current)!;
			const active = (element.getRootNode() as Document | ShadowRoot).activeElement;
			const label =
				active instanceof HTMLElement && content.contains(active)
					? active.getAttribute('aria-label')
					: null;
			handlers.setOption(key, value);
			queueMicrotask(() => {
				if (!element.open) return;
				draw(current);
				if (label)
					[...content.querySelectorAll<HTMLElement>('[aria-label]')]
						.find((control) => control.getAttribute('aria-label') === label)
						?.focus();
			});
		},
	};
	const draw = (page: BackstagePage) => {
		current = page;
		const content = pages.get(page)!;
		RENDERERS[page]({ handlers: pageHandlers, t, content } satisfies PageContext);
	};
	const localize = () => {
		const locale = normalizeEditorLocale(element.dataset.editorLocale);
		element.setAttribute('label', translate(locale, 'File'));
		element.setAttribute('back-label', translate(locale, 'Back to document'));
		element.items = NAV.map(({ id, label, footer }) => ({
			id,
			label: t(label),
			...(footer ? { group: 'footer' as const } : {}),
		}));
	};
	localize();

	element.addEventListener('office-backstage-select', (event) => {
		const { id } = (event as CustomEvent<{ id: string }>).detail;
		if (id === 'save') {
			event.preventDefault();
			handlers.close();
			handlers.fileCommand('save' satisfies FileCommand);
		} else draw(id as BackstagePage);
	});
	// Back and Escape: the editor decides what closing means (it also closes via `close()`).
	element.addEventListener('office-backstage-close', (event) => {
		if ((event as CustomEvent<{ reason: string }>).detail.reason === 'api') return;
		event.preventDefault();
		handlers.close();
	});
	return {
		element,
		open(page = 'home') {
			focus.open();
			draw(page);
			element.show(page);
		},
		close() {
			if (element.open) element.open = false;
			focus.close();
		},
		get isOpen() {
			return element.open;
		},
		relocalize() {
			localize();
			if (element.open) draw(current);
		},
	};
}
