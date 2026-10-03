/** Shared building blocks and the host interface of the File backstage pages. */
import type { WorkbookProperties } from '@christophervr/xlsx-core';
import type { EditorContext } from '../context';
import type { FileCommand } from '../events';
import { ribbonIcon } from '../ribbon/icons';
import type { SaveState } from '../title-bar';
import type { TemplateId } from './templates';

export type CalculationMode = 'automatic' | 'manual';
export interface EditorOptions {
	locale: string;
	theme: string;
	author: string;
	calculation: CalculationMode;
}

/** What the File view needs from the editor; every value is read when a page opens. */
export interface BackstageHost {
	readonly ctx: EditorContext;
	fileName(): string;
	saveState(): SaveState;
	fileCommand(command: FileCommand, fileName?: string): void;
	newFromTemplate(id: TemplateId): void;
	close(): void;
	options(): EditorOptions;
	setOption<K extends keyof EditorOptions>(key: K, value: EditorOptions[K]): void;
	/** Sets a workbook property (title, author, ...); an empty value clears it. */
	setProperty<K extends keyof WorkbookProperties>(key: K, value: WorkbookProperties[K]): void;
	passwordProtected?(): boolean;
	setPassword?(): Promise<void>;
	removePassword?(): void;
	hiddenActions(): readonly string[];
	setHiddenActions(ids: string[]): void;
}

export interface PageContext {
	host: BackstageHost;
	content: HTMLElement;
	t: EditorContext['t'];
	doc: Document;
}

export function heading(page: PageContext, text: string, level: 'h2' | 'h3' = 'h2'): HTMLElement {
	const node = page.doc.createElement(level);
	node.textContent = text;
	return node;
}

export function paragraph(page: PageContext, text: string, className = ''): HTMLElement {
	const node = page.doc.createElement('p');
	node.textContent = text;
	if (className) node.className = className;
	return node;
}

export function primary(page: PageContext, label: string, run: () => void): HTMLButtonElement {
	const button = page.doc.createElement('button');
	button.type = 'button';
	button.className = 'xve-backstage-primary';
	button.textContent = label;
	button.addEventListener('click', run);
	return button;
}

export function labelled(page: PageContext, label: string, control: HTMLElement): HTMLElement {
	const wrap = page.doc.createElement('label');
	wrap.className = 'xve-backstage-field';
	const text = page.doc.createElement('span');
	text.textContent = label;
	control.setAttribute('aria-label', label);
	wrap.append(text, control);
	return wrap;
}

/** An action card: a title, a description and one button. */
export function card(
	page: PageContext,
	title: string,
	description: string,
	action: HTMLElement,
): HTMLElement {
	const node = page.doc.createElement('div');
	node.className = 'xve-backstage-card';
	const text = page.doc.createElement('div');
	text.append(heading(page, title, 'h3'), paragraph(page, description, 'xve-backstage-muted'));
	node.append(text, action);
	return node;
}

export function facts(page: PageContext, rows: Array<[string, string]>): HTMLElement {
	const list = page.doc.createElement('dl');
	list.className = 'xve-backstage-facts';
	for (const [label, value] of rows) {
		const term = page.doc.createElement('dt');
		term.textContent = page.t(label);
		const detail = page.doc.createElement('dd');
		detail.textContent = value;
		list.append(term, detail);
	}
	return list;
}

export function noteList(page: PageContext, notes: string[], empty: string): HTMLElement {
	const list = page.doc.createElement('ul');
	list.className = 'xve-compatibility-notes';
	for (const note of notes) {
		const item = page.doc.createElement('li');
		const text = page.doc.createElement('span');
		text.textContent = note;
		item.append(ribbonIcon(page.doc, 'warning', 16), text);
		item.querySelector('svg')?.classList.add('xve-icon');
		list.append(item);
	}
	if (!notes.length) {
		const item = page.doc.createElement('li');
		item.textContent = empty;
		list.append(item);
	}
	return list;
}

/** Runs a file command after closing the backstage, as Excel does. */
export const runFile = (page: PageContext, command: FileCommand, fileName?: string) => () => {
	page.host.close();
	page.host.fileCommand(command, fileName);
};
