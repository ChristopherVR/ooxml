/** PowerPoint's Table grid: hover (or arrow) to size, click (or Enter) to insert. */
export const TABLE_PICKER_COLUMNS = 10;
export const TABLE_PICKER_ROWS = 8;

/** The grid's styling; the Insert tab and the phone Insert sheet both draw the picker. */
export const TABLE_PICKER_CSS = `
/* Table grid picker: 10 columns by 8 rows, the hovered block highlighted. */
.table-picker { padding:6px 10px 8px; }
.table-picker-label { min-height:18px; margin-bottom:4px; font-size:12px; }
.table-picker-grid { display:flex; flex-direction:column; gap:2px; width:max-content; outline-offset:2px; }
.table-picker-grid:focus-visible { outline:2px solid var(--pptx-ring,#6366f1); }
.table-picker-row { display:flex; gap:2px; }
.table-picker-cell { box-sizing:border-box; width:18px; height:18px; border:1px solid var(--pptx-border,#374151); border-radius:2px; }
.table-picker-cell[data-on] { background:var(--pptx-accent,#33334d); border-color:var(--pptx-primary,#6366f1); }
.table-picker-grid[aria-disabled] { opacity:.35; pointer-events:none; }
@media (pointer:coarse),(max-width:900px) {
	.table-picker-cell { width:26px; height:26px; }
}
`;

export interface TablePicker {
	el: HTMLElement;
	/** The text labels; `size` formats the hovered dimensions. */
	sync(labels: { heading: string; size: (columns: number, rows: number) => string }): void;
	/** Clear the highlight, e.g. when the menu opens. */
	reset(): void;
	disabled: boolean;
}

export function createTablePicker(
	doc: Document,
	pick: (rows: number, columns: number) => void,
): TablePicker {
	const el = doc.createElement('div');
	el.className = 'table-picker';
	const heading = doc.createElement('div');
	heading.className = 'table-picker-label';
	heading.setAttribute('aria-live', 'polite');
	const grid = doc.createElement('div');
	grid.className = 'table-picker-grid';
	grid.setAttribute('role', 'grid');
	grid.tabIndex = 0;
	// The menu moves focus to this stop with the other entries.
	grid.dataset.menuStop = '';
	const cells: HTMLElement[] = [];
	for (let row = 1; row <= TABLE_PICKER_ROWS; row += 1) {
		const line = doc.createElement('div');
		line.setAttribute('role', 'row');
		line.className = 'table-picker-row';
		for (let column = 1; column <= TABLE_PICKER_COLUMNS; column += 1) {
			const cell = doc.createElement('div');
			cell.setAttribute('role', 'gridcell');
			cell.className = 'table-picker-cell';
			cell.dataset.row = String(row);
			cell.dataset.column = String(column);
			cells.push(cell);
			line.append(cell);
		}
		grid.append(line);
	}
	el.append(heading, grid);
	let rows = 0;
	let columns = 0;
	let size: (columns: number, rows: number) => string = (c, r) => `${c}x${r}`;
	let title = '';
	const show = (nextRows: number, nextColumns: number) => {
		rows = nextRows;
		columns = nextColumns;
		for (const cell of cells) {
			cell.toggleAttribute(
				'data-on',
				Number(cell.dataset.row) <= rows && Number(cell.dataset.column) <= columns,
			);
		}
		heading.textContent = rows > 0 ? size(columns, rows) : title;
	};
	grid.addEventListener('pointerover', (event) => {
		const cell = (event.target as Element).closest<HTMLElement>('[data-row]');
		if (cell) {
			show(Number(cell.dataset.row), Number(cell.dataset.column));
		}
	});
	grid.addEventListener('pointerleave', () => show(0, 0));
	grid.addEventListener('click', (event) => {
		const cell = (event.target as Element).closest<HTMLElement>('[data-row]');
		if (cell && !picker.disabled) {
			pick(Number(cell.dataset.row), Number(cell.dataset.column));
		}
	});
	grid.addEventListener('focus', () => show(rows || 1, columns || 1));
	grid.addEventListener('blur', () => show(0, 0));
	grid.addEventListener('keydown', (event) => {
		const step: Record<string, [number, number]> = {
			ArrowRight: [0, 1],
			ArrowLeft: [0, -1],
			ArrowDown: [1, 0],
			ArrowUp: [-1, 0],
		};
		const move = step[event.key];
		if (move) {
			event.preventDefault();
			event.stopPropagation();
			show(
				Math.min(TABLE_PICKER_ROWS, Math.max(1, (rows || 1) + move[0])),
				Math.min(TABLE_PICKER_COLUMNS, Math.max(1, (columns || 1) + move[1])),
			);
		} else if ((event.key === 'Enter' || event.key === ' ') && rows > 0 && !picker.disabled) {
			event.preventDefault();
			event.stopPropagation();
			pick(rows, columns);
		}
	});
	const picker: TablePicker = {
		el,
		disabled: false,
		sync(labels) {
			title = labels.heading;
			size = labels.size;
			grid.setAttribute('aria-label', labels.heading);
			grid.toggleAttribute('aria-disabled', picker.disabled);
			show(rows, columns);
		},
		reset: () => show(0, 0),
	};
	show(0, 0);
	return picker;
}
