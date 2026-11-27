import type { EditorView } from 'prosemirror-view';
import type { EditorHost } from './editor-host';

export const DOCUMENT_VIEWS = [
	['draft', 'Draft'],
	['print', 'Print Layout'],
	['web', 'Web Layout'],
	['read', 'Read Mode'],
	['outline', 'Outline'],
	['focus', 'Focus'],
] as const;

export type DocumentView = (typeof DOCUMENT_VIEWS)[number][0];

export function isDocumentView(value: string): value is DocumentView {
	return DOCUMENT_VIEWS.some(([mode]) => mode === value);
}

export type OutlineAction =
	| { type: 'outline'; key: 'level' | 'paragraphLevel'; value: number }
	| {
			type: 'outline';
			key:
				| 'promote'
				| 'demote'
				| 'up'
				| 'down'
				| 'expand'
				| 'collapse'
				| 'format'
				| 'firstLine'
				| 'close';
	  };

export interface ViewHost extends EditorHost {
	canvas: HTMLElement;
	paper: HTMLElement;
	frame: HTMLElement;
	toolbar: HTMLElement | undefined;
	leave(): void;
	reviewAuthor?(): string;
	collaborating?(): boolean;
	revisionIds?(): ((kind: string) => string) | undefined;
}

export interface ViewPresentation {
	sync?(view: EditorView | undefined): void;
	run?(action: OutlineAction): void;
	/** Closes a layer the view owns; true when Escape was used. */
	escape?(): boolean;
	dispose(): void;
}

export function mountViewStyle(host: ViewHost, css: string): HTMLStyleElement {
	const style = host.element.ownerDocument.createElement('style');
	style.textContent = css;
	host.frame.append(style);
	return style;
}
