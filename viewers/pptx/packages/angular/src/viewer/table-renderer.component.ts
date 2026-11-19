import { NgStyle } from '@angular/common';
import {
	afterNextRender,
	ChangeDetectionStrategy,
	Component,
	DestroyRef,
	computed,
	effect,
	ElementRef,
	Injector,
	inject,
	input,
	output,
	signal,
	viewChild,
} from '@angular/core';
import type {
	TableCellCaret,
	TableCellEditEnd,
	TableCellTextEditorSession,
	TableStyleContext,
} from 'ooxml-ui/pptx';
import {
	canDrillDown,
	DEFAULT_FONT_FAMILY,
	openTableCellTextEditor,
	resolveTableCellEditEnd,
	screenStrokeWidth,
	tableCellEditorSegments,
	tableCellPointerIntent,
	tableContainerCss,
	tableFrameCss,
	linkedContentView,
} from 'ooxml-ui/pptx';
import type { PptxElement, PptxTableData, PptxTableRow, TablePptxElement } from 'pptx-viewer-core';

import { EditorStateService } from './editor-state.service';
import type { StyleMap } from './element-style';
import { injectLinkedContentAllowed, modelElementById } from './external-content.service';
import { LoadContentService } from './load-content.service';
import { SLIDE_CONTEXT } from './slide-context';
import {
	buildColStyles,
	buildTableViewModel,
	tableStyleContextFor,
} from './table-renderer-helpers';
import type { CellParagraph, TableRowViewModel } from './table-renderer-helpers';
import { TableResizeOverlayComponent } from './table-resize-overlay.component';
import { TableSelectionService } from './table-selection.service';
import { TABLE_STYLE_CONTEXT_OVERRIDE } from './table-style-context';

/** A committed table-cell text edit (original grid coordinates + new text). */
export interface TableCellCommit {
	rowIndex: number;
	colIndex: number;
	text: string;
	/** Set by the rich cell editor: why it closed, with the runs it committed. */
	end?: TableCellEditEnd;
}

/** The cell currently being edited (original grid coordinates), or null. */
interface EditingCell {
	rowIndex: number;
	colIndex: number;
	caret: TableCellCaret;
}

/**
 * May this table's individual cells be selected/edited at all?
 *
 * Pure (no DI, no signals) so it is unit-testable without a full Angular
 * injection context: `TableRendererComponent`'s constructor runs an
 * `effect()` that needs a `ChangeDetectionScheduler` this package's
 * TestBed-free suite doesn't provide (see `ribbon-home-section.component.ts`'s
 * `performResetSlide` for the same extraction pattern). G8: `a:
 * graphicFrameLocks/@noDrilldown` forbids selecting/editing this table's
 * individual cells, even on an otherwise-editable deck.
 */
export function canDrillDownIntoTable(editable: boolean, element: PptxElement): boolean {
	return editable && canDrillDown(element);
}

/**
 * Must a pointer-down on a cell stop before it reaches the slide canvas?
 *
 * The canvas selects on `pointerdown`, and a Shift press there TOGGLES the table
 * out of the element selection, which hides the Table Layout tab before the
 * cell's `click` can stretch the range (so Merge Cells was unreachable). The rule
 * is shared `tableCellPointerIntent`; an `extend` verdict consumes the press, as
 * Svelte's `applyTableCellPointer` and Vue's `onCellPointerDown` do.
 */
export function cellPressConsumesShift(input: {
	canDrillDown: boolean;
	shiftKey: boolean;
	elementSelected: boolean;
	rangeElementId: string | undefined;
	elementId: string;
}): boolean {
	return (
		input.canDrillDown &&
		tableCellPointerIntent({
			isTableCell: true,
			shiftKey: input.shiftKey,
			elementSelected: input.elementSelected,
			rangeOnSameElement: input.rangeElementId === input.elementId,
		}) === 'extend'
	);
}

/**
 * TableRendererComponent: Angular port of the React `renderTableFromTableData`
 * (packages/react/src/viewer/utils/table-render-data.tsx) plus the editor
 * overlays (`table-render.tsx` selection, `table-render-resize.tsx` handles).
 *
 * Renders a `<table>` from the typed `PptxTableData` structure. Behaviours:
 *  - Merged cells resolve to `colspan`/`rowspan`; banding is baked into the
 *    per-cell style; diagonal borders render as an SVG overlay.
 *  - Editing (when `editable`): single click selects a cell, Shift+Click extends
 *    a rectangular range (both via {@link TableSelectionService}, shared with the
 *    inspector), double click opens an inline text input, and drag handles on the
 *    column / row boundaries resize the table (emitted through `tableChange`).
 *
 * Pure helpers live in `table-renderer-helpers.ts` / `table-cell-style.ts`.
 */
@Component({
	selector: 'pptx-table-renderer',
	standalone: true,
	changeDetection: ChangeDetectionStrategy.OnPush,
	imports: [NgStyle, TableResizeOverlayComponent],
	templateUrl: './table-renderer.component.html',
	styleUrl: './table-renderer.component.css',
})
export class TableRendererComponent {
	/**
	 * Font stack for table text with no authored typeface, declared on the
	 * `<table>` root. Load-bearing: an unstyled cell otherwise inherits the HOST
	 * chrome's font, so the same table measured a different stack in every
	 * binding's demo. All five bindings declare this same shared default.
	 */
	readonly defaultTableFontFamily = DEFAULT_FONT_FAMILY;
	/** Diagonal border width on screen: never under one device pixel. */
	protected readonly screenStrokeWidth = screenStrokeWidth;
	private readonly injector = inject(Injector);
	/** Shared cell-selection state (present only inside the editor subtree). */
	private readonly selectionSvc = inject(TableSelectionService, { optional: true });
	/** The slide's element selection (present only inside the editor subtree). */
	private readonly editorState = inject(EditorStateService, { optional: true });
	/**
	 * Deck load state (present in the viewer subtree). Supplies the parsed
	 * table-style map + theme colour/font scheme so table-style-inherited banding,
	 * fonts, and section diagonals resolve. Absent (optional) in standalone
	 * thumbnail/export renderers, where the hardcoded defaults are used.
	 */
	private readonly loader = inject(LoadContentService, { optional: true });
	/** A surface's own table style context (the thumbnail rail's), preferred over the loader's. */
	private readonly styleOverride = inject(TABLE_STYLE_CONTEXT_OVERRIDE, { optional: true });
	private readonly linkedContentAllowed = injectLinkedContentAllowed();
	private readonly slideContext = inject(SLIDE_CONTEXT, { optional: true });

	/** The table element to render. Must be `type === 'table'`. */
	readonly element = input.required<PptxElement>();

	/** Whether inline cell editing / selection / resize is enabled. */
	readonly editable = input<boolean>(false);

	/** Emitted when a cell edit is committed (Enter / blur). */
	readonly cellCommit = output<TableCellCommit>();

	/** Emitted when a structural table change (resize) should be persisted. */
	readonly tableChange = output<{ id: string; tableData: PptxTableData }>();

	/** The cell currently being edited (original grid coords), or null. */
	private readonly editingCell = signal<EditingCell | null>(null);

	/** The host the rich cell editor mounts into while a cell is edited. */
	private readonly cellEditorHost = viewChild<ElementRef<HTMLDivElement>>('cellEditorHost');

	/** The open rich cell editor, if any. */
	private cellSession: TableCellTextEditorSession | undefined;

	/**
	 * `<table>`-level style beyond the font stack: `a:tblPr@rtl` mirrors the
	 * column order for right-to-left decks. Decided in shared so all five
	 * bindings honour the flag identically.
	 */
	readonly tableRootStyle = computed<StyleMap>(() => {
		const el = this.element();
		const data = el.type === 'table' ? el.tableData : undefined;
		return tableContainerCss(data) as StyleMap;
	});

	/** Pre-computed `<col>` styles for the colgroup. */
	readonly colStyles = computed<StyleMap[]>(() => buildColStyles(this.element()));

	/**
	 * Table-style context (parsed style map + theme colour/font scheme) from the
	 * deck loader, or `undefined` outside the viewer subtree. `fontScheme` lets a
	 * table style's `a:fontRef@idx` resolve to a concrete font family.
	 */
	private readonly styleCtx = computed<TableStyleContext | undefined>(() => {
		const context = this.styleOverride
			? this.styleOverride()
			: this.loader
				? tableStyleContextFor(this.loader.tableStyleMap(), this.loader.theme())
				: undefined;
		// A table style's picture fill is drawn only once this viewer allows linked content.
		return linkedContentView(context, this.linkedContentAllowed());
	});

	/** The table style's drop shadow (Themed Style 1/2), drawn on the table's frame. */
	readonly tableFrameStyle = computed<StyleMap>(
		() => tableFrameCss(this.tableData(), this.styleCtx()) as StyleMap,
	);

	/** Projected view-model rows with merged-cell resolution + banding applied. */
	readonly rows = computed<TableRowViewModel[]>(() =>
		buildTableViewModel(this.element(), this.styleCtx()),
	);

	/** Column widths (0-1 fractions) for the resize overlay. */
	readonly columnWidths = computed<number[]>(() => this.tableData()?.columnWidths ?? []);

	/** The table's rows, so the resize overlay skips boundaries inside merged cells. */
	readonly tableRows = computed<PptxTableRow[]>(() => this.tableData()?.rows ?? []);

	constructor() {
		// Mount the shared rich editor once the edited cell's host renders; close
		// it (committing if still open) when editing ends or moves elsewhere.
		effect(() => {
			const cell = this.editingCell();
			if (!cell) {
				const open = this.cellSession;
				this.cellSession = undefined;
				open?.destroy();
				return;
			}
			afterNextRender(() => this.openCellEditor(cell), { injector: this.injector });
		});
		inject(DestroyRef).onDestroy(() => {
			const open = this.cellSession;
			this.cellSession = undefined;
			open?.destroy();
		});
	}

	private tableData(): PptxTableData | undefined {
		const el = this.element();
		return el.type === 'table' ? (el as TablePptxElement).tableData : undefined;
	}

	/** The model's table data, for edits: the drawn element may withhold linked content. */
	private modelTableData(): PptxTableData | undefined {
		const found = modelElementById(
			this.editorState,
			this.element().id,
			this.slideContext?.slideId() ?? null,
		);
		return found?.element.type === 'table' ? found.element.tableData : this.tableData();
	}

	/** True when the given cell (original grid coords) is being edited. */
	isEditing(rowIndex: number, colIndex: number): boolean {
		const e = this.editingCell();
		return e !== null && e.rowIndex === rowIndex && e.colIndex === colIndex;
	}

	/** True when the cell is the current selection anchor for this element. */
	isSelectedAnchor(rowIndex: number, colIndex: number): boolean {
		const sel = this.selectionSvc?.selection();
		return (
			Boolean(sel) &&
			sel?.elementId === this.element().id &&
			sel?.rowIndex === rowIndex &&
			sel?.columnIndex === colIndex
		);
	}

	/** True when the cell falls inside the active Shift+Click range. */
	isInRange(rowIndex: number, colIndex: number): boolean {
		const sel = this.selectionSvc?.selection();
		if (!sel || sel.elementId !== this.element().id || !sel.selectedCells) {
			return false;
		}
		return sel.selectedCells.some((c) => c.row === rowIndex && c.col === colIndex);
	}

	/** A Shift press that extends the range never reaches the canvas's element toggle. */
	onCellPointerDown(event: PointerEvent): void {
		const id = this.element().id;
		if (
			cellPressConsumesShift({
				canDrillDown:
					Boolean(this.selectionSvc) && canDrillDownIntoTable(this.editable(), this.element()),
				shiftKey: event.shiftKey,
				elementSelected: this.editorState?.isSelected(id) ?? false,
				rangeElementId: this.selectionSvc?.selection()?.elementId,
				elementId: id,
			})
		) {
			event.stopPropagation();
		}
	}

	/** Single click selects the cell; Shift+Click extends a rectangular range. */
	onCellClick(event: MouseEvent, rowIndex: number, colIndex: number): void {
		if (!canDrillDownIntoTable(this.editable(), this.element()) || !this.selectionSvc) {
			return;
		}
		event.stopPropagation();
		const id = this.element().id;
		const td = this.tableData();
		if (event.shiftKey && td) {
			this.selectionSvc.extendTo(id, rowIndex, colIndex, td);
		} else {
			this.selectionSvc.selectCell(id, rowIndex, colIndex);
		}
	}

	/** Double-click on a cell enters inline edit mode. */
	onCellDblClick(event: Event, rowIndex: number, colIndex: number): void {
		if (!canDrillDownIntoTable(this.editable(), this.element())) {
			return;
		}
		event.stopPropagation();
		this.editingCell.set({ rowIndex, colIndex, caret: 'all' });
		this.selectionSvc?.beginEditing(this.element().id, rowIndex, colIndex);
	}

	/** Mount the shared rich cell editor into the edited cell's host. */
	private openCellEditor(cell: EditingCell): void {
		this.cellSession?.destroy();
		this.cellSession = undefined;
		const host = this.cellEditorHost()?.nativeElement;
		const data = this.tableData();
		if (!host || !this.isEditing(cell.rowIndex, cell.colIndex)) {
			return;
		}
		this.cellSession = openTableCellTextEditor({
			host,
			segments: tableCellEditorSegments(
				data?.rows[cell.rowIndex]?.cells[cell.colIndex],
				this.element() as TablePptxElement,
				cell.rowIndex,
				cell.colIndex,
			),
			caret: cell.caret,
			rtl: data?.rtl === true,
			onEnd: (end) => this.onCellEditEnd(cell, end),
		});
	}

	/**
	 * The rich cell editor closed. The shared `resolveTableCellEditEnd` outcome
	 * names the next cell (Tab, Shift+Tab, an arrow at the text's edge); the
	 * commit handler applies the same outcome to the document.
	 */
	onCellEditEnd(cell: EditingCell, end: TableCellEditEnd): void {
		const element = this.element();
		const outcome =
			element.type === 'table'
				? resolveTableCellEditEnd(element as TablePptxElement, cell.rowIndex, cell.colIndex, end)
				: {};
		const current = this.editingCell();
		if (current?.rowIndex === cell.rowIndex && current.colIndex === cell.colIndex) {
			if (outcome.edit) {
				this.editingCell.set({
					rowIndex: outcome.edit.row,
					colIndex: outcome.edit.col,
					caret: outcome.edit.caret,
				});
				this.selectionSvc?.beginEditing(element.id, outcome.edit.row, outcome.edit.col);
			} else {
				this.editingCell.set(null);
				this.selectionSvc?.endEditing();
				if (outcome.select) {
					this.selectionSvc?.selectCell(element.id, outcome.select.row, outcome.select.col);
				}
			}
		}
		this.cellCommit.emit({ rowIndex: cell.rowIndex, colIndex: cell.colIndex, text: '', end });
	}
	/** Persist a column-width drag by emitting new table data. */
	onResizeColumns(newWidths: number[]): void {
		const td = this.modelTableData();
		if (!td) {
			return;
		}
		this.tableChange.emit({
			id: this.element().id,
			tableData: { ...td, columnWidths: newWidths },
		});
	}

	/** Persist a row-height drag by emitting new table data. */
	onResizeRow(event: { index: number; height: number }): void {
		const td = this.modelTableData();
		if (!td) {
			return;
		}
		const rows = td.rows.map((row, i) =>
			i === event.index ? { ...row, height: event.height } : row,
		);
		this.tableChange.emit({ id: this.element().id, tableData: { ...td, rows } });
	}
}

// Re-export for template type-checking (CellParagraph is used in the @for loop).
export type { CellParagraph };
