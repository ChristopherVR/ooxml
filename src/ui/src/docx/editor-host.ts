import { isUpdateRefusal, UpdateRefusedError } from 'ooxml-core/collab';
import { translateCollabText } from 'ooxml-core/i18n';
import { emit } from './events';
import type { EditorView } from 'prosemirror-view';
import type { DocumentModel } from 'ooxml-core/docx';
import { normalizeEditorLocale, type EditorLocale } from './localization';

/** What feature controllers need from the editor element, without reaching into its internals. */
export interface EditorHost {
	readonly element: HTMLElement;
	view(): EditorView | undefined;
	model(): DocumentModel;
	/** Replaces the model for edits made outside the main editing surface. */
	setModel(model: DocumentModel): void;
	locale(): EditorLocale;
	/** Editable and not in a collaboration session (parts and page setup aren't synced). */
	canEditOutsideBody(): boolean;
	/** Marks the document changed after an edit outside the main surface. */
	edited(): void;
	reportError(cause: unknown): void;
}

/**
 * Raises the editor's `document-error` event for any thrown value. A refused collaboration update
 * says why sharing stopped in the editor's display language.
 */
export function dispatchDocumentError(element: HTMLElement, cause: unknown): void {
	const error = cause instanceof Error ? cause : new Error(String(cause));
	emit(
		element,
		'document-error',
		isUpdateRefusal(error)
			? new UpdateRefusedError(
					error.code,
					error.detail,
					translateCollabText(
						error.message,
						normalizeEditorLocale((element as { locale?: string }).locale),
					),
				)
			: error,
	);
}
