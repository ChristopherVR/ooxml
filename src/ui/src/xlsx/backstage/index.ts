/**
 * Excel's File view. The shared `office-ui-backstage` draws the navigation column (Back, Info, New,
 * Open, Save and its siblings, Options with Customize Ribbon) and shows one page at a time; each
 * page is a light-DOM child that this module renders with its own content.
 */
import { defineBackstage } from '../../controls';
import type { OfficeBackstageItem } from '../../controls';
import type { BackstageHost, PageContext } from './parts';
import { renderInfo } from './pages-info';
import { renderExport, renderNew, renderOpen, renderPrint, renderSaveAs } from './pages-file';
import { createExcelOptions } from './options-dialog';
import { renderShare } from './pages-share';

export type { BackstageHost, EditorOptions, CalculationMode, IterationSettings } from './parts';
export { participants, type ShareHost } from './pages-share';
export { compatibilityNotes, formatNotes } from './pages-info';
export { createTemplateWorkbook, TEMPLATES, type TemplateId } from 'ooxml-core/xlsx/ui';

/** The File view's pages; `options` and `customize` open the Options dialog instead. */
export type BackstagePage =
	| 'info'
	| 'new'
	| 'open'
	| 'saveAs'
	| 'print'
	| 'share'
	| 'export'
	| 'options'
	| 'customize';
type ContentPage = Exclude<BackstagePage, 'options' | 'customize'>;

export interface Backstage {
	readonly element: HTMLElement;
	/** File > Options: a dialog over the workbook, as PowerPoint and Visio show it. */
	readonly optionsDialog: HTMLElement;
	open(page?: BackstagePage): void;
	close(): void;
	readonly isOpen: boolean;
	/** Re-renders the open page (after a locale change). */
	relocalize(): void;
	/** Re-renders `page` when it is the one open (its state changed). */
	refresh(page: BackstagePage): void;
}

type BackstageElement = HTMLElement & {
	items: readonly OfficeBackstageItem[];
	open: boolean;
	show(page?: string): void;
	close(): void;
};

/** Pages in navigation order; Save has no page of its own and Options sit in the footer group. */
const PAGES: Array<[ContentPage, string]> = [
	['info', 'Info'],
	['new', 'New'],
	['open', 'Open'],
	['saveAs', 'Save As'],
	['print', 'Print'],
	['share', 'Share'],
	['export', 'Export'],
];
const RENDERERS: Record<ContentPage, (page: PageContext) => void> = {
	info: renderInfo,
	new: renderNew,
	open: renderOpen,
	saveAs: renderSaveAs,
	print: renderPrint,
	share: renderShare,
	export: renderExport,
};

export function createBackstage(host: BackstageHost): Backstage {
	const { ctx } = host;
	const doc = ctx.host.ownerDocument;
	defineBackstage();
	const element = doc.createElement('office-ui-backstage') as BackstageElement;
	element.className = 'xve-backstage';
	element.setAttribute('part', 'backstage');
	const contents = new Map<ContentPage, HTMLElement>();
	for (const [page] of PAGES) {
		const content = doc.createElement('div');
		content.className = 'xve-backstage-content';
		content.dataset.backstagePage = page;
		content.hidden = true;
		contents.set(page, content);
		element.append(content);
	}
	let current: ContentPage = 'info';
	const options = createExcelOptions(host);
	const show = (page: ContentPage) => {
		current = page;
		RENDERERS[page]({ host, content: contents.get(page)!, t: ctx.t, doc });
	};
	/** Options and Customize Ribbon are one dialog: the backstage steps aside for it. */
	const showOptions = (category?: string) => {
		if (element.open) element.close();
		host.close();
		options.show(category);
	};
	element.addEventListener('office-backstage-select', (event) => {
		const id = (event as CustomEvent<{ id: string }>).detail.id;
		if (id === 'save') {
			host.close();
			host.fileCommand('save');
		} else if (id === 'options') showOptions();
		else if (contents.has(id as ContentPage)) show(id as ContentPage);
	});
	// Back and Escape close the element itself; the host restores focus and tracks the state.
	element.addEventListener('office-backstage-close', (event) => {
		if ((event as CustomEvent<{ reason: string }>).detail.reason !== 'api') host.close();
	});
	const relocalize = () => {
		element.setAttribute('label', ctx.t('File'));
		element.setAttribute('back-label', ctx.t('Back to workbook'));
		const item = (id: string, label: string, group?: 'footer'): OfficeBackstageItem => ({
			id,
			label: ctx.t(label),
			...(group ? { group } : {}),
		});
		const label = new Map(PAGES);
		const page = (id: ContentPage) => item(id, label.get(id)!);
		element.items = [
			page('info'),
			page('new'),
			page('open'),
			item('save', 'Save'),
			page('saveAs'),
			page('print'),
			page('share'),
			page('export'),
			item('options', 'Options', 'footer'),
		];
		options.relocalize();
		if (element.open) show(current);
	};
	relocalize();
	return {
		element,
		optionsDialog: options.element,
		open(page = 'info') {
			if (page === 'options' || page === 'customize') {
				showOptions(page === 'customize' ? 'ribbon' : undefined);
				return;
			}
			show(page);
			element.show(page);
		},
		close() {
			element.close();
		},
		get isOpen() {
			return element.open;
		},
		relocalize,
		refresh(page) {
			if (element.open && current === page) show(page);
		},
	};
}
