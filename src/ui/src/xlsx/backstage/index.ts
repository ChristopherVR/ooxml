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
import { renderCustomize, renderOptions } from './pages-options';

export type { BackstageHost, EditorOptions, CalculationMode, IterationSettings } from './parts';
export { compatibilityNotes, formatNotes } from './pages-info';
export { createTemplateWorkbook, TEMPLATES, type TemplateId } from 'ooxml-core/xlsx/ui';

export type BackstagePage =
	| 'info'
	| 'new'
	| 'open'
	| 'saveAs'
	| 'print'
	| 'export'
	| 'options'
	| 'customize';

export interface Backstage {
	readonly element: HTMLElement;
	open(page?: BackstagePage): void;
	close(): void;
	readonly isOpen: boolean;
	/** Re-renders the open page (after a locale change). */
	relocalize(): void;
}

type BackstageElement = HTMLElement & {
	items: readonly OfficeBackstageItem[];
	open: boolean;
	show(page?: string): void;
	close(): void;
};

/** Pages in navigation order; Save has no page of its own and Options sit in the footer group. */
const PAGES: Array<[BackstagePage, string]> = [
	['info', 'Info'],
	['new', 'New'],
	['open', 'Open'],
	['saveAs', 'Save As'],
	['print', 'Print'],
	['export', 'Export'],
	['options', 'Options'],
	['customize', 'Customize Ribbon'],
];
const RENDERERS: Record<BackstagePage, (page: PageContext) => void> = {
	info: renderInfo,
	new: renderNew,
	open: renderOpen,
	saveAs: renderSaveAs,
	print: renderPrint,
	export: renderExport,
	options: renderOptions,
	customize: renderCustomize,
};

export function createBackstage(host: BackstageHost): Backstage {
	const { ctx } = host;
	const doc = ctx.host.ownerDocument;
	defineBackstage();
	const element = doc.createElement('office-ui-backstage') as BackstageElement;
	element.className = 'xve-backstage';
	element.setAttribute('part', 'backstage');
	const contents = new Map<BackstagePage, HTMLElement>();
	for (const [page] of PAGES) {
		const content = doc.createElement('div');
		content.className = 'xve-backstage-content';
		content.dataset.backstagePage = page;
		content.hidden = true;
		contents.set(page, content);
		element.append(content);
	}
	let current: BackstagePage = 'info';
	// A changed language or theme re-renders the open page so its text follows.
	const pageHost: BackstageHost = {
		...host,
		ctx,
		setOption(key, value) {
			host.setOption(key, value);
			queueMicrotask(() => {
				if (element.open) show(current);
			});
		},
	};
	const show = (page: BackstagePage) => {
		current = page;
		RENDERERS[page]({ host: pageHost, content: contents.get(page)!, t: ctx.t, doc });
	};
	element.addEventListener('office-backstage-select', (event) => {
		const id = (event as CustomEvent<{ id: string }>).detail.id;
		if (id === 'save') {
			host.close();
			host.fileCommand('save');
		} else if (contents.has(id as BackstagePage)) show(id as BackstagePage);
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
		const page = (id: BackstagePage, group?: 'footer') => item(id, label.get(id)!, group);
		element.items = [
			page('info'),
			page('new'),
			page('open'),
			item('save', 'Save'),
			page('saveAs'),
			page('print'),
			page('export'),
			page('options', 'footer'),
			page('customize', 'footer'),
		];
		if (element.open) show(current);
	};
	relocalize();
	return {
		element,
		open(page = 'info') {
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
	};
}
