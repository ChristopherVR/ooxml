import {
	columnHandleSegments,
	computeColumnBoundaries,
	computeResizedColumnWidths,
	computeResizedRowHeight,
	computeTableMergeCrossings,
	DEFAULT_ROW_HEIGHT,
	rowHandleSegments,
} from 'ooxml-ui/pptx';
import type { PptxTableRow } from 'pptx-viewer-core';
import React, { useRef, useEffect, useMemo, useLayoutEffect, useState, useCallback } from 'react';

/**
 * Overlay that renders draggable column and row resize handles on top of a table.
 * A boundary is drawn only along real cell edges: the stretch inside a merged
 * cell gets no handle, so it never covers (and steals clicks from) that cell.
 */
export function TableResizeOverlay({
	children,
	columnWidths,
	rows,
	editable,
	onResizeColumns,
	onResizeRow,
}: {
	children: React.ReactNode;
	/** Column widths as proportions summing to ~1 */
	columnWidths: number[];
	/** The table's rows, for the merges that interrupt a boundary. */
	rows?: readonly PptxTableRow[];
	editable: boolean;
	onResizeColumns?: (newWidths: number[]) => void;
	onResizeRow?: (rowIndex: number, newHeight: number) => void;
}) {
	const containerRef = useRef<HTMLDivElement>(null);
	const [rowBounds, setRowBounds] = useState<number[]>([]);
	const [tableHeight, setTableHeight] = useState(0);

	// Drag state stored in a ref to avoid re-renders mid-drag
	const dragRef = useRef<{
		type: 'col' | 'row';
		index: number;
		startPos: number;
		handleEl: HTMLDivElement;
		initialWidths?: number[];
		initialRowHeight?: number;
	} | null>(null);

	// Column boundary positions (cumulative percentages)
	const colBoundaries = useMemo(() => computeColumnBoundaries(columnWidths), [columnWidths]);
	const crossings = useMemo(
		() => computeTableMergeCrossings(rows ?? [], columnWidths.length),
		[rows, columnWidths.length],
	);

	// Measure row boundaries after layout
	const measureRows = useCallback(() => {
		const container = containerRef.current;
		if (!container) {
			return;
		}
		const table = container.querySelector('table');
		if (!table) {
			return;
		}
		const trs = table.querySelectorAll('tbody > tr');
		const bounds: number[] = [];
		let cumHeight = 0;
		trs.forEach((tr, i) => {
			cumHeight += (tr as HTMLElement).offsetHeight;
			if (i < trs.length - 1) {
				bounds.push(cumHeight);
			}
		});
		setTableHeight(table.offsetHeight);
		// Only update state when bounds actually change to avoid infinite re-render loop
		setRowBounds((prev) => {
			if (prev.length === bounds.length && prev.every((v, i) => v === bounds[i])) {
				return prev;
			}
			return bounds;
		});
	}, []);

	// Re-measure whenever content changes
	useLayoutEffect(() => {
		measureRows();
	});

	// Global mouse handlers for drag
	useEffect(() => {
		const handleMouseMove = (e: MouseEvent) => {
			const drag = dragRef.current;
			if (!drag) {
				return;
			}
			e.preventDefault();
			const delta = drag.type === 'col' ? e.clientX - drag.startPos : e.clientY - drag.startPos;
			drag.handleEl.style.transform =
				drag.type === 'col' ? `translateX(${delta}px)` : `translateY(${delta}px)`;
		};

		const handleMouseUp = (e: MouseEvent) => {
			const drag = dragRef.current;
			if (!drag || !containerRef.current) {
				return;
			}

			const rect = containerRef.current.getBoundingClientRect();

			if (
				e.clientX !== drag.startPos &&
				drag.type === 'col' &&
				drag.initialWidths &&
				onResizeColumns
			) {
				const deltaProp = (e.clientX - drag.startPos) / rect.width;
				onResizeColumns(computeResizedColumnWidths(drag.initialWidths, drag.index, deltaProp));
			} else if (e.clientY !== drag.startPos && drag.type === 'row' && onResizeRow) {
				const deltaY = e.clientY - drag.startPos;
				const newHeight = computeResizedRowHeight(
					drag.initialRowHeight ?? DEFAULT_ROW_HEIGHT,
					deltaY,
				);
				onResizeRow(drag.index, newHeight);
			}

			drag.handleEl.style.transform = '';
			document.body.style.cursor = '';
			document.body.style.userSelect = '';
			dragRef.current = null;
		};

		document.addEventListener('mousemove', handleMouseMove);
		document.addEventListener('mouseup', handleMouseUp);
		return () => {
			document.removeEventListener('mousemove', handleMouseMove);
			document.removeEventListener('mouseup', handleMouseUp);
		};
	}, [onResizeColumns, onResizeRow]);

	if (!editable) {
		return children;
	}

	const startColDrag = (e: React.MouseEvent<HTMLDivElement>, index: number) => {
		e.preventDefault();
		e.stopPropagation();
		document.body.style.cursor = 'col-resize';
		document.body.style.userSelect = 'none';
		dragRef.current = {
			type: 'col',
			index,
			startPos: e.clientX,
			handleEl: (e.currentTarget.parentElement as HTMLDivElement | null) ?? e.currentTarget,
			initialWidths: [...columnWidths],
		};
	};

	const startRowDrag = (e: React.MouseEvent<HTMLDivElement>, index: number) => {
		e.preventDefault();
		e.stopPropagation();
		const table = containerRef.current?.querySelector('table');
		const tr = table?.querySelectorAll('tbody > tr')[index];
		const actualHeight = (tr as HTMLElement)?.offsetHeight ?? DEFAULT_ROW_HEIGHT;
		document.body.style.cursor = 'row-resize';
		document.body.style.userSelect = 'none';
		dragRef.current = {
			type: 'row',
			index,
			startPos: e.clientY,
			handleEl: (e.currentTarget.parentElement as HTMLDivElement | null) ?? e.currentTarget,
			initialRowHeight: actualHeight,
		};
	};

	return (
		<div ref={containerRef} className='relative w-full h-full'>
			{children}

			{/* Column resize handles: one line per boundary, a segment per real edge */}
			{colBoundaries.map((leftPct, i) => (
				<div
					key={`col-h-${i}`}
					className='absolute top-0 bottom-0 w-[6px] z-10 pointer-events-none'
					style={{ left: `calc(${leftPct}% - 3px)` }}
				>
					{columnHandleSegments(crossings.columns[i] ?? [], rowBounds, tableHeight).map(
						({ top, height }) => (
							<div
								key={top}
								className='absolute left-0 right-0 cursor-col-resize pointer-events-auto group'
								style={{ top: `${top}px`, height: tableHeight ? `${height}px` : '100%' }}
								onMouseDown={(e) => startColDrag(e, i)}
							>
								<div className='w-px h-full mx-auto bg-transparent group-hover:bg-blue-400/60 transition-colors' />
							</div>
						),
					)}
				</div>
			))}

			{/* Row resize handles: one line per boundary, a segment per real edge */}
			{rowBounds.map((topPx, i) => (
				<div
					key={`row-h-${i}`}
					className='absolute left-0 right-0 h-[6px] z-10 pointer-events-none'
					style={{ top: `${topPx - 3}px` }}
				>
					{rowHandleSegments(crossings.rows[i] ?? [], columnWidths).map(({ leftPct, widthPct }) => (
						<div
							key={leftPct}
							className='absolute top-0 bottom-0 cursor-row-resize pointer-events-auto group'
							style={{ left: `${leftPct}%`, width: `${widthPct}%` }}
							onMouseDown={(e) => startRowDrag(e, i)}
						>
							<div className='h-px w-full my-auto bg-transparent group-hover:bg-blue-400/60 transition-colors' />
						</div>
					))}
				</div>
			))}
		</div>
	);
}
