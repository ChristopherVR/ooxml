/**
 * table-resize-overlay.component.ts: draggable column / row resize handles.
 *
 * Selector: `pptx-table-resize-overlay`
 *
 * Angular port of the React `TableResizeOverlay`
 * (packages/react/src/viewer/utils/table-render-resize.tsx). It projects the
 * rendered `<table>` via `<ng-content>` and overlays thin draggable handles on
 * the internal column boundaries and row boundaries. Drag geometry is delegated
 * to `pptx-viewer-shared` (`computeColumnBoundaries` / `computeResizedColumnWidths`
 * / `computeResizedRowHeight`); this component only wires pointer events.
 *
 * A boundary stretch inside a merged cell gets no handle
 * (`render/table-resize-merge.ts`), so a press there reaches the cell.
 *
 * On drop it emits the new column-width array (`resizeColumns`) or the resized
 * row's index + height (`resizeRow`); the parent commits them through the editor
 * history path.
 */
import {
	afterNextRender,
	ChangeDetectionStrategy,
	Component,
	computed,
	effect,
	ElementRef,
	Injector,
	inject,
	input,
	output,
	signal,
} from '@angular/core';
import {
	columnHandleSegments,
	computeColumnBoundaries,
	computeResizedColumnWidths,
	computeResizedRowHeight,
	computeTableMergeCrossings,
	DEFAULT_ROW_HEIGHT,
	rowHandleSegments,
} from 'ooxml-ui/pptx';
import type { TableRowHandleSegment } from 'ooxml-ui/pptx';
import type { PptxTableRow } from 'pptx-viewer-core';

interface DragState {
	type: 'col' | 'row';
	index: number;
	startPos: number;
	handle: HTMLElement;
	initialWidths?: number[];
	initialRowHeight?: number;
}

@Component({
	selector: 'pptx-table-resize-overlay',
	standalone: true,
	changeDetection: ChangeDetectionStrategy.OnPush,
	template: `
		<div class="pptx-ng-tbl-resize" #container>
			<ng-content />
			@if (editable()) {
				@for (leftPct of colBoundaries(); track $index; let i = $index) {
					<div class="pptx-ng-tbl-resize__col" [style.left.%]="leftPct">
						@for (segment of colSegments(i); track segment.top) {
							<div
								class="pptx-ng-tbl-resize__segment pptx-ng-tbl-resize__col-segment"
								[style.top.px]="segment.top"
								[style.height]="tableHeight() ? segment.height + 'px' : '100%'"
								(pointerdown)="onColDown($event, i)"
							>
								<div class="pptx-ng-tbl-resize__col-line"></div>
							</div>
						}
					</div>
				}
				@for (topPx of rowBounds(); track $index; let i = $index) {
					<div class="pptx-ng-tbl-resize__row" [style.top.px]="topPx">
						@for (segment of rowSegments(i); track segment.leftPct) {
							<div
								class="pptx-ng-tbl-resize__segment pptx-ng-tbl-resize__row-segment"
								[style.left.%]="segment.leftPct"
								[style.width.%]="segment.widthPct"
								(pointerdown)="onRowDown($event, i)"
							>
								<div class="pptx-ng-tbl-resize__row-line"></div>
							</div>
						}
					</div>
				}
			}
		</div>
	`,
	styleUrl: './table-resize-overlay.component.css',
})
export class TableResizeOverlayComponent {
	/** Column widths as proportions summing to ~1. */
	readonly columnWidths = input.required<number[]>();
	/** Whether the resize handles are active. */
	readonly editable = input<boolean>(false);
	/** The table's rows, for the merges that interrupt a boundary. */
	readonly rows = input<readonly PptxTableRow[]>([]);

	/** Emitted on column-boundary drop with the renormalised width array. */
	readonly resizeColumns = output<number[]>();
	/** Emitted on row-boundary drop with the resized row's index + new height. */
	readonly resizeRow = output<{ index: number; height: number }>();

	private readonly host = inject(ElementRef) as ElementRef<HTMLElement>;
	private readonly injector = inject(Injector);

	/** Cumulative internal column-boundary positions as percentages (0-100). */
	readonly colBoundaries = computed<number[]>(() => computeColumnBoundaries(this.columnWidths()));

	/** Measured internal row-boundary offsets (px from the table top). */
	readonly rowBounds = signal<number[]>([]);
	/** Measured table height, for the column-boundary segments. */
	readonly tableHeight = signal(0);

	/** Where merged cells interrupt the internal boundaries. */
	readonly crossings = computed(() =>
		computeTableMergeCrossings(this.rows(), this.columnWidths().length),
	);

	/** The real-edge segments of column boundary `index`, in px from the table top. */
	colSegments(index: number): { top: number; height: number }[] {
		return columnHandleSegments(
			this.crossings().columns[index] ?? [],
			this.rowBounds(),
			this.tableHeight(),
		);
	}

	/** The real-edge segments of row boundary `index`, as percentages of the width. */
	rowSegments(index: number): TableRowHandleSegment[] {
		return rowHandleSegments(this.crossings().rows[index] ?? [], this.columnWidths());
	}

	private drag: DragState | null = null;
	private readonly onMove = (e: PointerEvent): void => this.handleMove(e);
	private readonly onUp = (e: PointerEvent): void => this.handleUp(e);

	constructor() {
		// Initial measure once the projected table has mounted, then keep the row
		// boundaries in sync: a ResizeObserver (when available) catches row-height
		// changes, and an effect re-measures when the column set changes.
		afterNextRender(
			() => {
				this.measureRows();
				this.observeResize();
			},
			{ injector: this.injector },
		);
		effect(() => {
			// Depend on the column widths so structural changes trigger a re-measure.
			this.columnWidths();
			this.rows();
			afterNextRender(() => this.measureRows(), { injector: this.injector });
		});
	}

	/** Observe the container so row-height changes re-measure the boundaries. */
	private observeResize(): void {
		const container = this.container();
		if (!container || typeof ResizeObserver === 'undefined') {
			return;
		}
		const observer = new ResizeObserver(() => this.measureRows());
		observer.observe(container);
	}

	private container(): HTMLElement | null {
		return this.host.nativeElement.querySelector('.pptx-ng-tbl-resize');
	}

	private measureRows(): void {
		const table = this.container()?.querySelector('table');
		if (!table) {
			return;
		}
		this.tableHeight.set(table.offsetHeight);
		const trs = table.querySelectorAll('tbody > tr');
		const bounds: number[] = [];
		let cumulative = 0;
		trs.forEach((tr, i) => {
			cumulative += (tr as HTMLElement).offsetHeight;
			if (i < trs.length - 1) {
				bounds.push(cumulative);
			}
		});
		const prev = this.rowBounds();
		if (prev.length !== bounds.length || prev.some((v, i) => v !== bounds[i])) {
			this.rowBounds.set(bounds);
		}
	}

	onColDown(event: PointerEvent, index: number): void {
		event.preventDefault();
		event.stopPropagation();
		this.beginDrag({
			type: 'col',
			index,
			startPos: event.clientX,
			handle: this.boundaryOf(event),
			initialWidths: [...this.columnWidths()],
		});
	}

	onRowDown(event: PointerEvent, index: number): void {
		event.preventDefault();
		event.stopPropagation();
		const table = this.container()?.querySelector('table');
		const tr = table?.querySelectorAll('tbody > tr')[index] as HTMLElement | undefined;
		this.beginDrag({
			type: 'row',
			index,
			startPos: event.clientY,
			handle: this.boundaryOf(event),
			initialRowHeight: tr?.offsetHeight ?? DEFAULT_ROW_HEIGHT,
		});
	}

	/** The whole boundary line a pressed segment belongs to, so the drag moves all of it. */
	private boundaryOf(event: PointerEvent): HTMLElement {
		const segment = event.currentTarget as HTMLElement;
		return segment.parentElement ?? segment;
	}

	private beginDrag(state: DragState): void {
		this.drag = state;
		document.addEventListener('pointermove', this.onMove);
		document.addEventListener('pointerup', this.onUp);
	}

	private handleMove(event: PointerEvent): void {
		const drag = this.drag;
		if (!drag) {
			return;
		}
		event.preventDefault();
		const delta =
			drag.type === 'col' ? event.clientX - drag.startPos : event.clientY - drag.startPos;
		drag.handle.style.transform =
			drag.type === 'col' ? `translateX(${delta}px)` : `translateY(${delta}px)`;
	}

	private handleUp(event: PointerEvent): void {
		const drag = this.drag;
		this.drag = null;
		document.removeEventListener('pointermove', this.onMove);
		document.removeEventListener('pointerup', this.onUp);
		if (!drag) {
			return;
		}
		drag.handle.style.transform = '';

		if (event.clientX !== drag.startPos && drag.type === 'col' && drag.initialWidths) {
			const rect = this.container()?.getBoundingClientRect();
			const width = rect?.width ?? 1;
			const deltaProp = (event.clientX - drag.startPos) / width;
			this.resizeColumns.emit(
				computeResizedColumnWidths(drag.initialWidths, drag.index, deltaProp),
			);
		} else if (event.clientY !== drag.startPos && drag.type === 'row') {
			const deltaY = event.clientY - drag.startPos;
			const height = computeResizedRowHeight(drag.initialRowHeight ?? DEFAULT_ROW_HEIGHT, deltaY);
			this.resizeRow.emit({ index: drag.index, height });
		}
	}
}
