/**
 * Excel's File view. Left: Back, Info, New, Open, Save and its siblings (Save As, Print, Export)
 * and Options with Customize Ribbon; right: the selected page. Mirrors docx-viewer's backstage.ts.
 */
import { ribbonIcon } from '../ribbon/icons';
import type { BackstageHost, PageContext } from './parts';
import { renderInfo } from './pages-info';
import { renderExport, renderNew, renderOpen, renderPrint, renderSaveAs } from './pages-file';
import { renderCustomize, renderOptions } from './pages-options';

export type { BackstageHost, EditorOptions, CalculationMode } from './parts';
export { compatibilityNotes, formatNotes } from './pages-info';
export { createTemplateWorkbook, TEMPLATES, type TemplateId } from './templates';

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

const PAGES: Array<[BackstagePage, string, string]> = [
	['info', 'info', 'Info'],
	['new', 'newFile', 'New'],
	['open', 'open', 'Open'],
	['saveAs', 'copy', 'Save As'],
	['print', 'print', 'Print'],
	['export', 'export', 'Export'],
	['options', 'settings', 'Options'],
	['customize', 'settings', 'Customize Ribbon'],
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
	const element = doc.createElement('section');
	element.className = 'xve-backstage';
	element.setAttribute('part', 'backstage');
	element.setAttribute('role', 'dialog');
	element.setAttribute('aria-modal', 'true');
	element.hidden = true;
	const nav = doc.createElement('nav');
	nav.className = 'xve-backstage-nav';
	const content = doc.createElement('div');
	content.className = 'xve-backstage-content';
	const navButton = (icon: string) => {
		const button = doc.createElement('button');
		button.type = 'button';
		button.className = 'xve-backstage-nav-item';
		button.append(ribbonIcon(doc, icon, 16), doc.createElement('span'));
		return button;
	};
	const setText = (button: HTMLButtonElement, text: string) => {
		button.querySelector('span')!.textContent = ctx.t(text);
	};
	let current: BackstagePage = 'info';
	const buttons = new Map<BackstagePage, HTMLButtonElement>();
	// A changed language or theme re-renders the open page so its text follows.
	const pageHost: BackstageHost = {
		...host,
		ctx,
		setOption(key, value) {
			host.setOption(key, value);
			queueMicrotask(() => {
				if (!element.hidden) show(current);
			});
		},
	};
	const show = (page: BackstagePage) => {
		current = page;
		for (const [key, button] of buttons) button.setAttribute('aria-current', String(key === page));
		RENDERERS[page]({ host: pageHost, content, t: ctx.t, doc });
	};
	const back = navButton('back');
	back.classList.add('xve-backstage-back');
	back.addEventListener('click', () => host.close());
	const save = navButton('save');
	save.addEventListener('click', () => {
		host.close();
		host.fileCommand('save');
	});
	const separator = () => {
		const line = doc.createElement('div');
		line.className = 'xve-backstage-separator';
		line.setAttribute('role', 'separator');
		return line;
	};
	nav.append(back);
	for (const [page, icon] of PAGES) {
		const button = navButton(icon);
		button.addEventListener('click', () => show(page));
		buttons.set(page, button);
		nav.append(button);
		if (page === 'open') nav.append(separator(), save);
		if (page === 'export') nav.append(separator());
	}
	element.addEventListener('keydown', (event) => {
		if (event.key === 'Escape') {
			event.stopPropagation();
			host.close();
		}
	});
	const relocalize = () => {
		element.setAttribute('aria-label', ctx.t('File'));
		setText(back, 'Back to workbook');
		setText(save, 'Save');
		for (const [page, , label] of PAGES) setText(buttons.get(page)!, label);
		if (!element.hidden) show(current);
	};
	relocalize();
	element.append(nav, content);
	return {
		element,
		open(page = 'info') {
			element.hidden = false;
			show(page);
			back.focus();
		},
		close() {
			element.hidden = true;
		},
		get isOpen() {
			return !element.hidden;
		},
		relocalize,
	};
}
