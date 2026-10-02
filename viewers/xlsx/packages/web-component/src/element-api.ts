/**
 * The public methods of `<xlsx-editor>` (load, newWorkbook, save, saveBytes, download, markClean,
 * select, getSelection, setActiveSheet, undo, redo, focusGrid). The properties and lifecycle are
 * in component.ts; the state is the EditorCore.
 */
import { EditorCore } from './editor-core';
import { loadInto, newInto, saveBlob, saveBytes, type FileChrome } from './editor-files';
import { downloadBytes, saveExtension, withExtension } from './file-commands';
import { parseSelectionRef, selectionRef } from './selection';

/** Server rendering has no HTMLElement; the class still has to be definable there. */
export const HTMLElementBase = (
	typeof HTMLElement === 'undefined' ? class {} : HTMLElement
) as typeof HTMLElement;

export class XlsxEditorApi extends HTMLElementBase {
	protected readonly core: EditorCore = new EditorCore(this);
	/** Set by the shell once connected. */
	protected chrome: FileChrome | undefined;

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
