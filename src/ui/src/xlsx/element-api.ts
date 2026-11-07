import type { CloudOpenedFile, DocumentStorage } from 'ooxml-core/embedding';
/**
 * The public methods of `<xlsx-editor>` (load, newWorkbook, save, saveBytes, download, markClean,
 * select, getSelection, setActiveSheet, collaboration, share, undo, redo, focusGrid). The properties and lifecycle are
 * in component.ts; the state is the EditorCore.
 */
import { EditorCore } from './editor-core';
import {
	loadInto,
	newInto,
	openCloudFile,
	saveBlob,
	saveBytes,
	type FileChrome,
} from './editor-files';
import { downloadBytes, saveExtension, withExtension } from './file-commands';
import { parseSelectionRef, selectionRef } from 'ooxml-core/xlsx/ui';
import type { XlsxCollaborationOptions, XlsxCollaborationState } from './collaboration-types';
import { spreadsheetToolbarProps, type SpreadsheetToolbarProps } from './building-blocks';
import type { SurfaceHostActions } from './surface-actions';
import type { GoogleFontsOption } from 'ooxml-core/text';

/** Server rendering has no HTMLElement; the class still has to be definable there. */
export const HTMLElementBase = (
	typeof HTMLElement === 'undefined' ? class {} : HTMLElement
) as typeof HTMLElement;

export class XlsxEditorApi extends HTMLElementBase {
	/** Set before mounting to omit application UI. */
	surfaceOnly = false;
	getToolbarProps(): SpreadsheetToolbarProps {
		return spreadsheetToolbarProps(this.core);
	}
	/**
	 * Host callbacks that stand in for dialogs when `surfaceOnly` is set. A workflow without a
	 * callback is reported as not enabled in `getToolbarProps().enabled`.
	 */
	get surfaceActions(): SurfaceHostActions {
		return this.core.surfaceActions;
	}
	set surfaceActions(value: SurfaceHostActions | undefined) {
		this.core.surfaceActions = { ...(value ?? {}) };
		this.core.requestRender();
	}
	/**
	 * Host-owned storage. When set, File > Save and Save As hand the saved workbook to
	 * `storage.save()` instead of downloading it, and File > Open asks `storage.open()` (when it has
	 * one) instead of showing the file picker. The cancelable `file-command` event still fires
	 * first; a host that prevents it handles the command itself. Export and Save as CSV still
	 * download.
	 */
	get storage(): DocumentStorage | undefined {
		return this.core.storage;
	}
	set storage(value: DocumentStorage | null | undefined) {
		this.core.storage = value ?? undefined;
	}
	/**
	 * Webfont fallback for workbook fonts the reader has neither installed nor embedded. Unset (the
	 * default): the reader decides with File > Options > Trust Center > "Download fonts that are
	 * not installed or embedded", which is off until they turn it on. Set, it is the host's
	 * decision and that toggle is shown read-only: `false` is off (no request leaves the page;
	 * text uses installed or substitute fonts), `{ baseUrl }` requests the fonts the workbook uses
	 * that the bundled Google Fonts catalogue serves (or their metric-compatible clone, such as
	 * Carlito for Calibri) from that CSS2 endpoint (`https://fonts.googleapis.com/css2` for Google,
	 * or a self-hosted mirror) with one `<link>` in the document head, then the font files it
	 * names. Only public catalogue family names are sent, never cell text; an invalid `baseUrl`
	 * warns and stays off. Changing it applies to the open workbook.
	 */
	get googleFonts(): GoogleFontsOption | undefined {
		return this.core.webfonts.option;
	}
	set googleFonts(value: GoogleFontsOption | null | undefined) {
		this.core.webfonts.option = value ?? undefined;
		this.core.webfonts.apply();
	}
	get loading(): boolean {
		return this.core.loading;
	}
	subscribeEditorState(listener: () => void): () => void {
		this.addEventListener('editor-state-change', listener);
		return () => this.removeEventListener('editor-state-change', listener);
	}
	protected readonly core: EditorCore = new EditorCore(this);
	/** Set by the shell once connected. */
	protected chrome: FileChrome | undefined;
	/** Set by the shell once connected: opens the Share dialog. */
	protected openShare: (() => void) | undefined;
	/** Set by the shell once connected: leaves the File backstage. */
	protected closeBackstage: (() => void) | undefined;

	/** Opens .xlsx, .xlsm, .xltx, .xls or .csv bytes; `fileName` helps detection and names the file. */
	async load(bytes: Uint8Array | ArrayBuffer, fileName?: string): Promise<void> {
		await loadInto(this.core, bytes, fileName, this.chrome);
	}

	/**
	 * Opens a file a host got from a `CloudFileProvider` (its own picker or a start page), exactly as
	 * File > Open > OneDrive does: the discard confirmation when there are unsaved changes, the file's
	 * name, read-only when `file.canEdit` is false, and Save and Save As routed to `file.storage`.
	 * Resolves once open (or when the user declines to discard); rejects when the load fails.
	 */
	async openCloudFile(file: CloudOpenedFile): Promise<void> {
		await openCloudFile(this.core, file, this.chrome, () => this.closeBackstage?.());
	}

	/** Starts a blank workbook named Book1.xlsx. */
	newWorkbook(): void {
		newInto(this.core, this.chrome);
	}

	/** The workbook as .xlsx bytes (or the active sheet as CSV). Does not clear `dirty`. */
	saveBytes(format: 'xlsx' | 'csv' = 'xlsx'): Promise<Uint8Array> {
		return saveBytes(this.core, format);
	}

	/** The workbook as an .xlsx Blob. Does not clear `dirty`; call `markClean()` after persisting. */
	save(): Promise<Blob> {
		return saveBlob(this.core);
	}

	/** Saves and downloads in the browser, then marks the workbook clean. */
	async download(name?: string): Promise<void> {
		const fileName = name ?? withExtension(this.core.fileName, saveExtension(this.core.fileName));
		downloadBytes(this.ownerDocument, await this.saveBytes('xlsx'), fileName);
		this.markClean();
	}

	/** Clears the unsaved-changes flag (after the host persisted the bytes). */
	markClean(): void {
		this.core.dirty.set(false);
		this.chrome?.setSaveState('saved');
	}

	/** Commits an in-progress cell edit before a host navigates or persists. False means validation blocked it. */
	commitEdit(): boolean {
		const grid = this.core.ctx.grid();
		return !grid?.isEditing() || grid.commitEdit();
	}

	/** Selects `B2`, `B2:C5`, `A:A`, `B2:C5,E1` or `'Sheet 2'!A1` and scrolls it into view. */
	select(ref: string): void {
		if (!this.commitEdit()) throw new Error('The current cell edit failed validation.');
		const parsed = parseSelectionRef(ref, this.core.workbook);
		if (!parsed) throw new Error(`Not a cell reference: ${ref}`);
		if (parsed.sheet !== undefined) this.core.setActiveSheet(parsed.sheet);
		const first = parsed.ranges[0]!;
		this.core.selection.set({
			sheet: this.core.activeSheet,
			active: { ...first.start },
			anchor: { ...first.start },
			ranges: parsed.ranges,
		});
		this.core.ctx.grid()?.scrollTo(first.start);
	}

	/** The selection as A1 text (`B2:C5`, several ranges joined by commas). */
	getSelection(): string {
		return selectionRef(this.core.selection.get());
	}

	setActiveSheet(index: number): void {
		this.core.setActiveSheet(index);
	}

	/**
	 * Shares the open workbook in a room (see `XlsxCollaborationOptions`). Resolves once joined;
	 * the room's workbook replaces this one when the room already has one. Rejects when sharing is
	 * already active or no workbook is open.
	 */
	startCollaboration(options: XlsxCollaborationOptions): Promise<void> {
		return this.core.collab.start(options);
	}

	/** Leaves the room (and clears `collaboration`). The workbook stays as it is. */
	stopCollaboration(): void {
		this.core.collab.stop();
	}

	/** Restarts the connection, keeping the shared workbook and changes made offline. */
	reconnectCollaboration(): void {
		this.core.collab.reconnect();
	}

	/** Who is in the room and how the connection is. */
	get collaborationState(): XlsxCollaborationState {
		return this.core.collab.state();
	}

	/**
	 * Opens the Share dialog, where the user starts, joins or stops sharing, copies the invitation
	 * link and sees who is in the room. Without application UI (`surfaceOnly`) it calls the host's
	 * `onOpenShareDialog` instead, and does nothing when there is none (`getToolbarProps().enabled.openShareDialog` tells which).
	 */
	share(): void {
		if (this.surfaceOnly) void this.getToolbarProps().onOpenShareDialog();
		else this.openShare?.();
	}

	/** Undo the last edit; while shared, only this window's edits are undone. */
	undo(): void {
		void this.core.commands.run('edit.undo');
	}

	redo(): void {
		void this.core.commands.run('edit.redo');
	}

	focusGrid(): void {
		const grid = this.core.ctx.grid();
		if (grid) grid.focus();
		else this.shadowRoot?.querySelector<HTMLElement>('[part~="grid"]')?.focus();
	}
}
