import type {
	ExcelFileCommand,
	FileCommandDetail as SharedFileCommandDetail,
} from 'ooxml-core/embedding';
import type { RibbonAddInCommandDetail } from 'ooxml-core/ribbon';
import type { Workbook } from 'ooxml-core/xlsx';
import type { XlsxCollaborationState } from './collaboration-types';

/** The File commands this editor announces (the shared `ooxml-core/embedding` type). */
export type FileCommand = ExcelFileCommand;

/** The detail of `file-command`; the shared `ooxml-core/embedding` shape. */
export type FileCommandDetail = SharedFileCommandDetail<FileCommand>;

export interface SelectionChangeDetail {
	sheet: number;
	/** The selection as A1 text (`B2:C5`, or `A1:B2,D4` for several ranges). */
	ref: string;
	/** The active cell (`B2`). */
	active: string;
}

/** A cell edit that data validation (or the host's commit) refused. */
export interface ValidationErrorDetail {
	/** The message shown to the user (the rule's own text, or a generic one). */
	message: string;
	title?: string;
	/** The cell being edited, as A1 text. */
	ref: string;
	style?: 'stop' | 'warning' | 'information';
}

/** Every public event of `<xlsx-editor>` and its detail. All bubble and are composed. */
export interface XlsxEditorEventMap {
	/** The workbook was loaded, created or edited. */
	'workbook-change': { workbook: Workbook };
	/** Loading, saving, printing or a command failed. */
	'workbook-error': { error: Error; message: string };
	/** Non-fatal notes from the loader (unsupported features, compatibility). */
	'workbook-warning': { warnings: string[] };
	/** The user toggled Editing / Viewing, or the host changed `readOnly`. */
	'readonly-change': { readOnly: boolean };
	/** A command ran from the ribbon, a menu, Tell me or a shortcut. */
	'ribbon-action': { id: string };
	/** The user changed which ribbon commands are shown (File > Options > Customize Ribbon). */
	'ribbon-customize': { hiddenActions: string[] };
	/** Cancelable: call `preventDefault()` to handle a File command yourself. */
	'file-command': FileCommandDetail;
	'selection-change': SelectionChangeDetail;
	'sheet-change': { index: number; name: string };
	/** Unsaved-changes state flipped: true after an edit, false after save, load or `markClean()`. */
	'dirty-change': { dirty: boolean };
	/** Sharing started or stopped, the connection changed, or someone joined, left or synced. */
	'collaboration-change': XlsxCollaborationState;
	/** A command of a host tab (`ribbonAddIns`) was chosen. */
	'office-ribbon-add-in': RibbonAddInCommandDetail;
	/** An in-progress cell edit failed validation; it stays open until corrected or canceled. */
	'validation-error': ValidationErrorDetail;
}
export type XlsxEditorEventName = keyof XlsxEditorEventMap;
export type XlsxEditorEventDetail<K extends XlsxEditorEventName> = XlsxEditorEventMap[K];

/** Runtime list of every event name; the `satisfies` check keeps it a subset of the map. */
export const XLSX_EDITOR_EVENTS = [
	'workbook-change',
	'workbook-error',
	'workbook-warning',
	'readonly-change',
	'ribbon-action',
	'ribbon-customize',
	'file-command',
	'selection-change',
	'sheet-change',
	'dirty-change',
	'collaboration-change',
	'office-ribbon-add-in',
	'validation-error',
] as const satisfies readonly XlsxEditorEventName[];

/** Events a host can cancel with `preventDefault()`. */
const CANCELABLE = new Set<string>(['file-command']);

/**
 * Dispatches a bubbling, composed CustomEvent. Returns false when a cancelable event was canceled
 * (the `EventTarget.dispatchEvent` contract).
 */
export function emitEvent(target: EventTarget, type: string, detail: unknown): boolean {
	if (typeof CustomEvent === 'undefined') return true;
	return target.dispatchEvent(
		new CustomEvent(type, {
			detail,
			bubbles: true,
			composed: true,
			cancelable: CANCELABLE.has(type),
		}),
	);
}

/** Typed form of `emitEvent` for the public events. */
export function emit<K extends XlsxEditorEventName>(
	target: EventTarget,
	type: K,
	detail: XlsxEditorEventDetail<K>,
): boolean {
	return emitEvent(target, type, detail);
}
