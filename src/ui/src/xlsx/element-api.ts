/**
 * The public methods of `<xlsx-editor>` (load, newWorkbook, save, saveBytes, download, markClean,
 * select, getSelection, setActiveSheet, collaboration, share, undo, redo, focusGrid). The properties and lifecycle are
 * in component.ts; the state is the EditorCore.
 */
import { EditorCore } from './editor-core';
import { loadInto, newInto, saveBlob, saveBytes, type FileChrome } from './editor-files';
import { downloadBytes, saveExtension, withExtension } from './file-commands';
import { parseSelectionRef, selectionRef } from 'ooxml-core/xlsx/ui';
import type { XlsxCollaborationOptions, XlsxCollaborationState } from './collaboration-types';

/** Server rendering has no HTMLElement; the class still has to be definable there. */
export const HTMLElementBase = (
	typeof HTMLElement === 'undefined' ? class {} : HTMLElement
) as typeof HTMLElement;

export class XlsxEditorApi extends HTMLElementBase {
	protected readonly core: EditorCore = new EditorCore(this);
	/** Set by the shell once connected. */
	protected chrome: FileChrome | undefined;
	/** Set by the shell once connected: opens File > Share. */
	protected openShare: (() => void) | undefined;

	/** Opens .xlsx, .xlsm, .xltx, .xls or .csv bytes; `fileName` helps detection and names the file. */
	async load(bytes: Uint8Array | ArrayBuffer, fileName?: string): Promise<void> {
		await loadInto(this.core, bytes, fileName, this.chrome);
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

	/** Opens File > Share, where the user starts or stops sharing and sees who is in the room. */
	share(): void {
		this.openShare?.();
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
