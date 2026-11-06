/**
 * Maps the table-background gradient onto the table as it is laid out.
 *
 * `spanTableBackground` slices the gradient by the rows' declared heights, but
 * a row the text outgrows is taller than declared, so every slice below it
 * starts in the wrong place. The laid-out heights only exist once the table is
 * in the document, so a cell that shows the table gradient carries the
 * `--pptx-tblbg-span` marker and this module, armed once per page like
 * `paragraph-baseline-settle`, rewrites its `background-size` and
 * `background-position` in px from the real cell and table boxes. It works on
 * the rendered DOM, so every binding gets it without wiring of its own.
 *
 * @module render/table-style-bg-settle
 */

/** The custom property that marks a cell whose gradient slice is placed from the laid-out table. */
export const TABLE_BG_SPAN_VAR = '--pptx-tblbg-span';

/** A cell's box inside its table (layout px, unaffected by CSS transforms). */
export interface BoxInTable {
	left: number;
	top: number;
	width: number;
	height: number;
}

/**
 * The `background-size` and `background-position` (px) that show the part of
 * a table-sized gradient lying under `cell`. `layers` is how many background
 * layers the cell has: layers above the table gradient keep their own box.
 */
export function tableBackgroundSlice(
	cell: BoxInTable,
	table: { width: number; height: number },
	layers = 1,
): { size: string; position: string } {
	const round = (value: number) => Math.round(value * 100) / 100;
	const size = `${round(table.width)}px ${round(table.height)}px`;
	const position = `${round(-cell.left)}px ${round(-cell.top)}px`;
	return layers > 1
		? {
				size: `${'auto, '.repeat(layers - 1)}${size}`,
				position: `${'0 0, '.repeat(layers - 1)}${position}`,
			}
		: { size, position };
}

let armed = false;
let resizeObserver: ResizeObserver | undefined;
let scheduled = false;

/** Re-place every marked cell under `root`; returns how many were placed. */
export function settleTableBackgrounds(root: ParentNode = document): number {
	let placed = 0;
	for (const cell of root.querySelectorAll<HTMLElement>(`[style*="${TABLE_BG_SPAN_VAR}"]`)) {
		const table = cell.closest('table');
		if (!table || table.offsetWidth <= 0 || table.offsetHeight <= 0) {
			continue;
		}
		// A table whose rows grow or shrink later (an edit, a font that loads) is placed again.
		if (typeof ResizeObserver === 'function') {
			resizeObserver ??= new ResizeObserver(schedule);
			resizeObserver.observe(cell);
			resizeObserver.observe(table);
		}
		// Layers above the table gradient (a translucent tint) are recorded by the marker.
		const layers = Number.parseInt(cell.style.getPropertyValue(TABLE_BG_SPAN_VAR), 10) || 1;
		const { size, position } = tableBackgroundSlice(
			{
				left: cell.offsetLeft,
				top: cell.offsetTop,
				width: cell.offsetWidth,
				height: cell.offsetHeight,
			},
			{ width: table.offsetWidth, height: table.offsetHeight },
			layers,
		);
		if (cell.style.backgroundSize !== size) {
			cell.style.backgroundSize = size;
		}
		if (cell.style.backgroundPosition !== position) {
			cell.style.backgroundPosition = position;
		}
		placed += 1;
	}
	return placed;
}

function schedule(): void {
	if (scheduled) {
		return;
	}
	scheduled = true;
	const run = () => {
		scheduled = false;
		settleTableBackgrounds();
	};
	if (typeof requestAnimationFrame === 'function') {
		requestAnimationFrame(run);
	} else {
		setTimeout(run, 0);
	}
}

/**
 * Start watching for new and changed marked cells (once per page). Call it
 * when a cell was marked; it costs nothing before that.
 */
export function armTableBackgroundSettle(): void {
	if (armed || typeof document === 'undefined') {
		return;
	}
	armed = true;
	if (typeof MutationObserver === 'function') {
		new MutationObserver(schedule).observe(document.documentElement, {
			childList: true,
			subtree: true,
			characterData: true,
		});
	}
	schedule();
}
