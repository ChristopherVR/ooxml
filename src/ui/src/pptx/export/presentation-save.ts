/**
 * Where a saved presentation goes: the host's storage when the viewer has one (`storage` prop,
 * see `ooxml-core/embedding`), else a browser download. Every binding's Save / Save As path calls
 * this once, so routing a save to the host is decided in one place.
 */
import {
	documentMimeType,
	documentSaveMeta,
	saveToStorage,
	type DocumentSaveReason,
	type DocumentStorage,
} from 'ooxml-core/embedding';

import type { CompatibilityWarningToast } from '../render/compatibility-warning-toasts';
import { downloadBlob } from './download-helpers';

export type { DocumentSaveMeta, DocumentSaveReason, DocumentStorage } from 'ooxml-core/embedding';

/** The formats a viewer saves (`SavedPresentationFormat` of the render layer). */
export type PresentationSaveFormat = 'pptx' | 'ppsx' | 'pptm' | 'ppt';

/**
 * The MIME type a saved presentation is offered under, by format: legacy `.ppt` is an OLE2
 * compound file with its own type, and `.ppsx` / `.pptm` have theirs (`documentMimeType`).
 */
export function presentationSaveMimeType(format: PresentationSaveFormat): string {
	return documentMimeType(`presentation.${format}`);
}

export interface DeliverSavedPresentationInput {
	/** The host's storage; undefined (or null) downloads the file. */
	storage?: DocumentStorage | null | undefined;
	bytes: Uint8Array;
	format: PresentationSaveFormat;
	/** The name to save under, normally `savedPresentationFileName(source, format)`. */
	fileName: string;
	/** The command that saved: File > Save (`save`) or a Save As entry (`saveAs`), whatever the format. */
	reason: DocumentSaveReason;
}

/**
 * Hands the saved bytes to `storage.save()` when the viewer has a storage, else starts a browser
 * download. Resolves with the name the document was saved under (the one the host reported, if
 * any). Rejects when the storage does; a download never rejects.
 */
export async function deliverSavedPresentation(
	input: DeliverSavedPresentationInput,
): Promise<string> {
	const { storage, bytes, format, fileName } = input;
	const mimeType = presentationSaveMimeType(format);
	if (storage) {
		return saveToStorage(storage, bytes, documentSaveMeta(fileName, input.reason, mimeType));
	}
	downloadBlob(new Blob([bytes as BlobPart], { type: mimeType }), fileName);
	return fileName;
}

/** Toast code of a failed Save / Save As (the host's storage rejected, or serialising failed). */
export const SAVE_FAILED_TOAST_CODE = 'SAVE_FAILED';

/** The notice every binding shows when a Save / Save As fails, in its compatibility toast stack. */
export const SAVE_FAILED_TOAST: CompatibilityWarningToast = {
	id: 'save-failed',
	code: SAVE_FAILED_TOAST_CODE,
	severity: 'warning',
	messageKey: 'pptx.autosave.saveFailed',
};

/** The message a binding reports to the host's `error` event for a failed save. */
export function saveFailureMessage(error: unknown): string {
	const detail = error instanceof Error ? error.message : String(error);
	return detail ? `Save failed: ${detail}` : 'Save failed';
}

/** `toasts` with the save-failure notice first (once), for a binding's toast stack. */
export function withSaveFailureToast(
	toasts: readonly CompatibilityWarningToast[],
): CompatibilityWarningToast[] {
	return [SAVE_FAILED_TOAST, ...toasts.filter((toast) => toast.id !== SAVE_FAILED_TOAST.id)];
}
