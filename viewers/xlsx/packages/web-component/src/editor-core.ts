/**
 * The editor's state and the `EditorContext` every UI module receives: the workbook and its edit
 * session, the active sheet and selection, the command and dialog registries, and the public
 * event plumbing. The element (component.ts) owns one EditorCore; the shell (editor-shell.ts)
 * builds the DOM around it.
 */
import {
	approximateMeasure,
	createEditSession,
	type EditSession,
	type FontView,
	type Workbook,
	type Worksheet,
} from '@christophervr/xlsx-core';
import { createCommandRegistry, type CommandRegistry } from './commands';
import type { EditorContext, GridController, SelectionModel } from './context';
import { createDialogRegistry, type DialogRegistry } from './dialogs';
import { DirtyState } from './dirty-state';
import { emit, emitEvent } from './events';
import { DEFAULT_AUTHOR_NAME, DEFAULT_FILE_NAME } from './editor-attributes';
import {
	normalizeEditorLocale,
	sheetBaseName,
	translator,
	type EditorLocale,
} from './localization';
import {
	activeRef,
	createSelectionModel,
	initialSelection,
	selectionRef,
	SheetSelections,
} from './selection';
import type { CalculationMode } from './backstage';
import type { EditorThemeMode, XlsxTheme } from './theme';

/** What the shell exposes back to the core (filled in when the element connects). */
export interface ShellHooks {
	render(): void;
	toast(message: string, kind?: 'info' | 'warning' | 'error'): void;
	relocalize(): void;
}

export class EditorCore {
	workbook: Workbook | undefined;
	session: EditSession | undefined;
	activeSheet = 0;
	readOnly = false;
	locale: EditorLocale = 'en';
	authorName = DEFAULT_AUTHOR_NAME;
	fileName = DEFAULT_FILE_NAME;
	theme: EditorThemeMode = 'auto';
	showToolbar = true;
	showFormulaBar = true;
	themeColors: Partial<XlsxTheme> = {};
	appliedThemeVars: string[] = [];
	hiddenActions: string[] = [];
	loadGeneration = 0;
	savePassword: string | undefined;
	shell: ShellHooks | undefined;
	readonly selection: SelectionModel = createSelectionModel();
	readonly commands: CommandRegistry;
	readonly dialogs: DialogRegistry;
	readonly dirty: DirtyState;
	readonly ctx: EditorContext;
	private gridController: GridController | undefined;
	private readonly modelListeners = new Set<(change: unknown) => void>();
	private stopSession: (() => void) | undefined;
	private renderQueued = false;
	private lastSelection = '';
	private readonly sheetSelections = new SheetSelections();
	/** The worksheet object on screen, to notice sheets inserted, moved or deleted around it. */
	private shownSheet: Worksheet | undefined;

	/** Automatic or manual calculation: the workbook's own mode (`calcPr calcMode`). */
	get calculation(): CalculationMode {
		return this.workbook?.calcMode === 'manual' ? 'manual' : 'automatic';
	}

	constructor(readonly host: HTMLElement) {
		this.dirty = new DirtyState((dirty) => emit(host, 'dirty-change', { dirty }));
		const report = (what: string, error: unknown) => {
			const message = error instanceof Error ? error.message : String(error);
			this.shell?.toast(
				this.ctx.t('{what} failed: {message}', {
					what: this.ctx.t(what),
					message: this.ctx.t(message),
				}),
				'error',
			);
			const failure = error instanceof Error ? error : new Error(message);
			emit(host, 'workbook-error', { error: failure, message });
		};
		this.commands = createCommandRegistry(() => this.ctx, {
			onError: (id, error) => report(this.commands.get(id)?.label ?? id, error),
			isHidden: (id) => this.hiddenActions.includes(id),
		});
		this.dialogs = createDialogRegistry(
			() => this.ctx,
			(name, error) => report(name, error),
		);
		this.selection.onChange((selection) => {
			if (selection.sheet === this.activeSheet)
				this.sheetSelections.remember(this.shownSheet, selection);
			const detail = {
				sheet: selection.sheet,
				ref: selectionRef(selection),
				active: activeRef(selection),
			};
			this.ctx.emit('selection-change', detail);
			this.requestRender();
		});
		this.ctx = this.createContext();
	}

	private createContext(): EditorContext {
		const core = this;
		return {
			host: core.host,
			get root() {
				return core.host.shadowRoot!;
			},
			session: () => core.session,
			workbook: () => core.workbook,
			activeSheet: () => core.activeSheet,
			setActiveSheet: (index) => core.setActiveSheet(index),
			selection: core.selection,
			readOnly: () => core.readOnly,
			t: translator(() => core.locale),
			commands: core.commands,
			dialogs: core.dialogs,
			grid: () => core.gridController,
			attachGrid: (grid) => {
				core.gridController = grid;
				core.requestRender();
			},
			onModelChange: (listener) => {
				core.modelListeners.add(listener);
				return () => {
					core.modelListeners.delete(listener);
				};
			},
			requestRender: () => core.requestRender(),
			toast: (message, kind) => core.shell?.toast(message, kind),
			emit: (type, detail) => {
				// The grid and the selection model may both report one move; announce it once.
				if (type === 'selection-change') {
					const key = JSON.stringify(detail);
					if (key === core.lastSelection) return true;
					core.lastSelection = key;
				}
				return emitEvent(core.host, type, detail);
			},
			authorName: () => core.authorName,
			locale: () => core.locale,
		};
	}

	/** Debounced refresh of the ribbon, title bar and status bar. */
	requestRender(): void {
		if (this.renderQueued) return;
		this.renderQueued = true;
		queueMicrotask(() => {
			this.renderQueued = false;
			this.shell?.render();
		});
	}

	notifyModel(change: unknown): void {
		for (const listener of [...this.modelListeners]) {
			try {
				listener(change);
			} catch (error) {
				console.error(error);
			}
		}
		try {
			this.gridController?.invalidate();
		} catch {
			// A failing repaint must not break the edit that triggered it.
		}
		this.requestRender();
	}

	/** Replaces the workbook (load, new, template, property), with a fresh edit session. */
	setWorkbook(workbook: Workbook | undefined): void {
		this.savePassword = undefined;
		this.stopSession?.();
		this.stopSession = undefined;
		this.workbook = workbook;
		this.session = workbook ? this.createSession(workbook) : undefined;
		if (this.session && workbook?.fullCalcOnLoad && this.session.autoRecalc()) {
			try {
				this.session.calc.recalculateAll();
			} catch (error) {
				console.error(error);
			}
		}
		this.stopSession = this.session?.onChange((change) => this.onSessionChange(change));
		const sheets = workbook?.sheets.length ?? 0;
		const preferred = workbook?.activeSheet ?? 0;
		this.activeSheet = Math.min(Math.max(0, preferred), Math.max(0, sheets - 1));
		if (workbook?.sheets[this.activeSheet]?.state !== 'visible') {
			const visible = workbook?.sheets.findIndex((sheet) => sheet.state === 'visible') ?? -1;
			if (visible >= 0) this.activeSheet = visible;
		}
		this.lastSelection = '';
		this.shownSheet = workbook?.sheets[this.activeSheet];
		this.selection.set(initialSelection(this.activeSheet, workbook));
		this.dirty.set(false);
		this.notifyModel({ kind: 'load' });
	}

	private createSession(workbook: Workbook): EditSession | undefined {
		try {
			return createEditSession(workbook, {
				defaultSheetBase: sheetBaseName(this.ctx.t),
				autoRowHeight: true,
				measureText: (text, font) => this.measureText(text, font),
			});
		} catch (error) {
			// A model the session cannot drive (for example a partial object) is shown read-only.
			console.warn('xlsx-editor: the workbook could not be opened for editing', error);
			return undefined;
		}
	}

	private onSessionChange(change: unknown): void {
		const workbook = this.workbook;
		if (workbook) this.followShownSheet(workbook);
		this.dirty.set(true);
		if (workbook) emit(this.host, 'workbook-change', { workbook });
		this.notifyModel(change);
	}

	setActiveSheet(index: number, force = false): void {
		const sheet = this.workbook?.sheets[index];
		if (!sheet || (index === this.activeSheet && !force)) return;
		if (this.gridController?.isEditing()) this.gridController.commitEdit();
		this.activeSheet = index;
		this.shownSheet = sheet;
		this.selection.set(this.sheetSelections.restore(index, this.workbook));
		emit(this.host, 'sheet-change', { index, name: sheet.name });
		this.notifyModel({ kind: 'sheet', sheet: index });
	}

	/**
	 * After an edit that inserted, moved or deleted sheets: keep showing the same worksheet at its
	 * new index, or, when it was deleted, switch to the nearest visible sheet (always announcing
	 * `sheet-change` and restoring that sheet's selection, even when the index is unchanged).
	 */
	private followShownSheet(workbook: Workbook): void {
		const shown = this.shownSheet;
		if (workbook.sheets[this.activeSheet] === shown && shown) return;
		const moved = shown ? workbook.sheets.indexOf(shown) : -1;
		if (moved >= 0 && shown) {
			this.activeSheet = moved;
			this.selection.set({ ...this.selection.get(), sheet: moved });
			emit(this.host, 'sheet-change', { index: moved, name: shown.name });
			return;
		}
		const sheets = workbook.sheets;
		const from = Math.min(this.activeSheet, Math.max(0, sheets.length - 1));
		const after = sheets.findIndex((sheet, i) => i >= from && sheet.state === 'visible');
		const any = sheets.findIndex((sheet) => sheet.state === 'visible');
		const next = after >= 0 ? after : any >= 0 ? any : 0;
		if (sheets[next]) this.setActiveSheet(next, true);
		else this.shownSheet = undefined;
	}

	/** Text width for the core's automatic row heights, from the grid's canvas measurer. */
	private measureText(text: string, font: FontView): number {
		const grid = this.gridController;
		if (!grid) return approximateMeasure(text, font);
		const css = `${font.italic ? 'italic ' : ''}${font.bold ? 'bold ' : ''}${font.sizePx}px "${font.family}"`;
		return grid.measureText(text, css);
	}

	/** Switches automatic or manual calculation on the live session (one undo step, saved). */
	setCalculation(mode: CalculationMode): void {
		if (mode === this.calculation) return;
		this.session?.setCalcMode(mode === 'manual' ? 'manual' : 'auto');
	}

	setLocale(value: string): void {
		const next = normalizeEditorLocale(value);
		if (next === this.locale) return;
		this.locale = next;
		this.shell?.relocalize();
		this.notifyModel({ kind: 'locale' });
	}

	setReadOnly(readOnly: boolean): boolean {
		if (readOnly === this.readOnly) return false;
		this.readOnly = readOnly;
		if (readOnly && this.gridController?.isEditing()) this.gridController.cancelEdit();
		this.notifyModel({ kind: 'readonly', readOnly });
		emit(this.host, 'readonly-change', { readOnly });
		return true;
	}
}
