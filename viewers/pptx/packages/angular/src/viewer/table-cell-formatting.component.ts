/**
 * table-cell-formatting.component.ts: per-cell formatting inspector for the
 * currently selected table cell.
 *
 * Selector: `pptx-table-cell-formatting`
 *
 * Angular port of the React `TableCellFormattingPanel`. Reads the selected cell
 * from {@link TableSelectionService} (shared with the canvas renderer) and, when
 * that selection targets THIS table element, exposes font size, text/background
 * colour, advanced fill (via `pptx-table-cell-advanced-fill`), bold / italic /
 * underline, horizontal & vertical alignment, per-edge borders, and merge /
 * split actions. Every edit emits a fully-updated element through
 * `elementChange`, which the inspector commits as one undoable history entry.
 */
import { ChangeDetectionStrategy, Component, computed, inject, input, output } from '@angular/core';
import { TranslatePipe } from '@ngx-translate/core';
import { TABLE_CELL_SECTIONS, withCellStyle } from 'ooxml-ui/pptx';
import type { ThemeColorPickerCommit } from 'ooxml-ui/pptx';
import type { PptxTableCellStyle, PptxTableData, TablePptxElement } from 'pptx-viewer-core';

import { InspectorFieldsComponent } from './inspector-fields.component';
import { TableCellAdvancedFillComponent } from './table-cell-advanced-fill.component';
import { TableCellColorFieldComponent } from './table-cell-color-field.component';
import {
	mergeDown,
	mergeRight,
	mergeSelection,
	patchTableData,
	splitCursorCell,
} from './table-data-helpers';
import type { TableCellSelection } from './table-selection.service';
import { TableSelectionService } from './table-selection.service';

@Component({
	selector: 'pptx-table-cell-formatting',
	standalone: true,
	changeDetection: ChangeDetectionStrategy.OnPush,
	imports: [
		InspectorFieldsComponent,
		TableCellAdvancedFillComponent,
		TableCellColorFieldComponent,
		TranslatePipe,
	],
	template: `
		@if (cell(); as c) {
			<div class="pptx-tcf">
				<div class="pptx-tcf__heading">
					{{
						'pptx.table.cell' | translate: { row: sel()!.rowIndex + 1, col: sel()!.columnIndex + 1 }
					}}
				</div>

				<div class="pptx-tcf__grid2">
					<pptx-table-cell-color-field
						[label]="'pptx.table.color' | translate"
						[value]="style().color"
						fallback="#000000"
						[selectedRef]="style().colorRef"
						[disabled]="!canEdit()"
						(commit)="onColorCommit('color', 'colorRef', $event)"
					/>
					<pptx-table-cell-color-field
						[label]="'pptx.table.background' | translate"
						[value]="style().backgroundColor"
						fallback="#ffffff"
						[selectedRef]="style().backgroundColorRef"
						[disabled]="!canEdit()"
						(commit)="onColorCommit('backgroundColor', 'backgroundColorRef', $event)"
					/>
				</div>

				<pptx-table-cell-advanced-fill
					[cellStyle]="style()"
					[canEdit]="canEdit()"
					(styleChange)="updateStyle($event)"
				/>

				<pptx-inspector-fields
					[sections]="cellSections"
					[target]="style()"
					[disabled]="!canEdit()"
					(patch)="updateStyle($event)"
				/>

				<div class="pptx-tcf__btns">
					<button
						type="button"
						class="pptx-tcf__btn"
						[disabled]="!canEdit()"
						(click)="onMergeRight()"
					>
						{{ 'pptx.table.mergeRight' | translate }}
					</button>
					<button
						type="button"
						class="pptx-tcf__btn"
						[disabled]="!canEdit()"
						(click)="onMergeDown()"
					>
						{{ 'pptx.table.mergeDown' | translate }}
					</button>
					<button type="button" class="pptx-tcf__btn" [disabled]="!canEdit()" (click)="onSplit()">
						{{ 'pptx.table.split' | translate }}
					</button>
					@if (hasRange()) {
						<button
							type="button"
							class="pptx-tcf__btn"
							[disabled]="!canEdit()"
							(click)="onMergeRange()"
						>
							{{ 'pptx.contextMenu.mergeSelectedCells' | translate }}
						</button>
					}
				</div>
			</div>
		}
	`,
	styles: `
		.pptx-tcf {
			display: flex;
			flex-direction: column;
			gap: 0.35rem;
		}
		.pptx-tcf__heading {
			font-size: 10px;
			font-weight: 600;
			text-transform: uppercase;
			letter-spacing: 0.05em;
			color: var(--pptx-inspector-muted, #888);
		}
		.pptx-tcf__field {
			display: flex;
			align-items: center;
			gap: 0.35rem;
		}
		.pptx-tcf__grid2 {
			display: grid;
			grid-template-columns: 1fr 1fr;
			gap: 0.3rem;
		}
		.pptx-tcf__lbl {
			font-size: 10px;
			color: var(--pptx-inspector-muted, #888);
		}
		.pptx-tcf__num {
			flex: 1;
			min-width: 0;
			background: var(--pptx-inspector-input-bg, #2d2d2d);
			border: 1px solid var(--pptx-inspector-border, #444);
			color: inherit;
			border-radius: 3px;
			padding: 2px 4px;
			font-size: 11px;
		}
		.pptx-tcf__color {
			width: 28px;
			height: 22px;
			padding: 0;
			border: 1px solid var(--pptx-inspector-border, #444);
			border-radius: 3px;
			background: transparent;
			cursor: pointer;
		}
		.pptx-tcf__btns {
			display: flex;
			flex-wrap: wrap;
			gap: 0.25rem;
		}
		.pptx-tcf__toggle,
		.pptx-tcf__btn {
			padding: 2px 8px;
			font-size: 11px;
			background: var(--pptx-inspector-input-bg, #2d2d2d);
			border: 1px solid var(--pptx-inspector-border, #444);
			color: inherit;
			border-radius: 3px;
			cursor: pointer;
		}
		.pptx-tcf__toggle.is-active {
			background: var(--pptx-inspector-accent, #2563eb);
			color: #fff;
			border-color: var(--pptx-inspector-accent, #2563eb);
		}
		.pptx-tcf__toggle:disabled,
		.pptx-tcf__btn:disabled {
			opacity: 0.4;
			cursor: not-allowed;
		}
	`,
})
export class TableCellFormattingComponent {
	/** The table element being edited. */
	readonly element = input.required<TablePptxElement>();
	/** Whether editing is enabled. */
	readonly canEdit = input<boolean>(true);
	/** Emits the fully-updated element after any edit. */
	readonly elementChange = output<TablePptxElement>();

	private readonly selection = inject(TableSelectionService, { optional: true });

	/** The shared cell descriptors: size, emphasis, alignment, margins, borders. */
	protected readonly cellSections = TABLE_CELL_SECTIONS;

	/** The selection when it targets THIS element, else null. */
	protected readonly sel = computed(() => {
		const s = this.selection?.selection();
		return s && s.elementId === this.element().id ? s : null;
	});

	/** The selected cell, or undefined when the selection is out of range. */
	protected readonly cell = computed(() => {
		const s = this.sel();
		const td = this.element().tableData;
		return s && td ? td.rows[s.rowIndex]?.cells[s.columnIndex] : undefined;
	});

	protected readonly style = computed<PptxTableCellStyle>(() => this.cell()?.style ?? {});

	protected readonly hasRange = computed<boolean>(
		() => (this.sel()?.selectedCells?.length ?? 0) >= 2,
	);

	/**
	 * A `pptx-table-cell-color-field` commit (text colour or fill colour): sets
	 * both the hex and its ref field, so a theme-swatch pick keeps following
	 * the deck's theme after a later theme change, and a native pick clears any
	 * previously-stored ref (the field always emits `ref: undefined` for one).
	 */
	protected onColorCommit(
		hexKey: 'color' | 'backgroundColor',
		refKey: 'colorRef' | 'backgroundColorRef',
		commit: ThemeColorPickerCommit,
	): void {
		this.updateStyle({ [hexKey]: commit.hex, [refKey]: commit.ref });
	}

	/** Merge a style patch into the selected cell and emit the updated element. */
	protected updateStyle(patch: Partial<PptxTableCellStyle>): void {
		const s = this.sel();
		const td = this.element().tableData;
		if (!s || !td) {
			return;
		}
		const rows = td.rows.map((row, ri) =>
			ri !== s.rowIndex
				? row
				: {
						...row,
						cells: row.cells.map((c, ci) => (ci !== s.columnIndex ? c : withCellStyle(c, patch))),
					},
		);
		this.commit({ ...td, rows });
	}

	protected onMergeRight(): void {
		this.applyMerge((s) => mergeRight(this.element(), s.rowIndex, s.columnIndex));
	}

	protected onMergeDown(): void {
		this.applyMerge((s) => mergeDown(this.element(), s.rowIndex, s.columnIndex));
	}

	protected onSplit(): void {
		this.applyMerge((s) => splitCursorCell(this.element(), s.rowIndex, s.columnIndex));
	}

	protected onMergeRange(): void {
		const s = this.sel();
		if (!s?.selectedCells) {
			return;
		}
		this.elementChange.emit(mergeSelection(this.element(), s.selectedCells));
	}

	private applyMerge(op: (s: TableCellSelection) => TablePptxElement): void {
		const s = this.sel();
		if (!s) {
			return;
		}
		this.elementChange.emit(op(s));
	}

	private commit(tableData: PptxTableData): void {
		this.elementChange.emit(patchTableData(this.element(), tableData));
	}
}
