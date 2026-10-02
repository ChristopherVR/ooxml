// Pure context-menu models for the cell area, the row and column headers and the sheet tabs.
// Labels are English translation keys; the menu DOM translates them with `ctx.t`.

export interface MenuEntry {
	id: string;
	/** English key, translated with `ctx.t` when rendering. */
	label: string;
	command?: string;
	arg?: unknown;
	action?: () => void | Promise<void>;
	disabled?: boolean;
	separatorBefore?: boolean;
	shortcut?: string;
	checked?: boolean;
}

type Spec = Omit<MenuEntry, 'separatorBefore'> & { editing?: boolean };

/** Flattens groups, marks the first entry of each later group and applies read-only. */
function build(groups: Spec[][], readOnly: boolean): MenuEntry[] {
	const out: MenuEntry[] = [];
	groups
		.filter((group) => group.length > 0)
		.forEach((group, index) =>
			group.forEach((spec, position) => {
				const { editing = true, ...entry } = spec;
				const item: MenuEntry = { ...entry };
				if ((editing && readOnly) || spec.disabled) item.disabled = true;
				if (index > 0 && position === 0) item.separatorBefore = true;
				out.push(item);
			}),
		);
	return out;
}

const clipboardGroup = (): Spec[] => [
	{ id: 'cut', label: 'Cut', command: 'edit.cut', shortcut: 'Ctrl+X' },
	{ id: 'copy', label: 'Copy', command: 'edit.copy', shortcut: 'Ctrl+C', editing: false },
	{ id: 'paste', label: 'Paste', command: 'edit.paste', shortcut: 'Ctrl+V' },
	{ id: 'paste-special', label: 'Paste Special...', command: 'edit.paste-special' },
];

export interface CellMenuState {
	readOnly: boolean;
	hasComment: boolean;
	hasLink: boolean;
	hasList: boolean;
}

export function cellMenu(state: CellMenuState): MenuEntry[] {
	return build(
		[
			clipboardGroup(),
			[
				{ id: 'insert-cells', label: 'Insert...', command: 'cells.insert-cells' },
				{ id: 'delete-cells', label: 'Delete...', command: 'cells.delete-cells' },
				{ id: 'clear-contents', label: 'Clear Contents', command: 'edit.delete' },
			],
			[
				{ id: 'filter', label: 'Filter', command: 'data.filter' },
				{ id: 'sort-ascending', label: 'Sort A to Z', command: 'data.sort-ascending' },
				{ id: 'sort-descending', label: 'Sort Z to A', command: 'data.sort-descending' },
				{ id: 'custom-sort', label: 'Custom Sort...', command: 'data.custom-sort' },
			],
			state.hasComment
				? [
						{ id: 'edit-comment', label: 'Edit Comment', command: 'review.new-comment' },
						{ id: 'delete-comment', label: 'Delete Comment', command: 'review.delete-comment' },
					]
				: [{ id: 'new-comment', label: 'New Comment', command: 'review.new-comment' }],
			[
				{
					id: 'format-cells',
					label: 'Format Cells...',
					command: 'format.cells',
					shortcut: 'Ctrl+1',
				},
				{
					id: 'pick-from-list',
					label: 'Pick From Drop-down List...',
					command: 'grid.pick-from-list',
				},
				{ id: 'define-name', label: 'Define Name...', command: 'formulas.define-name' },
				...(state.hasLink
					? [
							{ id: 'edit-link', label: 'Edit Hyperlink', command: 'insert.link' },
							{ id: 'remove-link', label: 'Remove Hyperlink', command: 'insert.remove-link' },
						]
					: [{ id: 'link', label: 'Link', command: 'insert.link', shortcut: 'Ctrl+K' }]),
			],
		],
		state.readOnly,
	);
}

export interface HeaderMenuState {
	readOnly: boolean;
	hiddenInSelection: boolean;
}

function headerMenu(state: HeaderMenuState, axis: 'rows' | 'columns'): MenuEntry[] {
	const row = axis === 'rows';
	return build(
		[
			clipboardGroup(),
			[
				{
					id: 'insert',
					label: 'Insert',
					command: row ? 'cells.insert-rows' : 'cells.insert-columns',
				},
				{
					id: 'delete',
					label: 'Delete',
					command: row ? 'cells.delete-rows' : 'cells.delete-columns',
				},
				{ id: 'clear-contents', label: 'Clear Contents', command: 'edit.delete' },
			],
			[
				{
					id: 'format-cells',
					label: 'Format Cells...',
					command: 'format.cells',
					shortcut: 'Ctrl+1',
				},
				row
					? { id: 'row-height', label: 'Row Height...', command: 'format.row-height' }
					: { id: 'column-width', label: 'Column Width...', command: 'format.column-width' },
				{ id: 'hide', label: 'Hide', command: row ? 'format.hide-rows' : 'format.hide-columns' },
				{
					id: 'unhide',
					label: 'Unhide',
					command: row ? 'format.unhide-rows' : 'format.unhide-columns',
					disabled: !state.hiddenInSelection,
				},
			],
		],
		state.readOnly,
	);
}

export const rowHeaderMenu = (state: HeaderMenuState): MenuEntry[] => headerMenu(state, 'rows');
export const columnHeaderMenu = (state: HeaderMenuState): MenuEntry[] =>
	headerMenu(state, 'columns');

export interface TabMenuState {
	readOnly: boolean;
	structureLocked: boolean;
	visibleSheets: number;
	hiddenSheets: number;
	protected: boolean;
}

export function tabMenu(state: TabMenuState, actions: { rename(): void }): MenuEntry[] {
	const locked = state.structureLocked;
	const single = state.visibleSheets <= 1;
	return build(
		[
			[
				{ id: 'insert', label: 'Insert', command: 'sheet.insert', disabled: locked },
				{ id: 'delete', label: 'Delete', command: 'sheet.delete', disabled: locked || single },
				{ id: 'rename', label: 'Rename', action: () => actions.rename(), disabled: locked },
				{ id: 'move-copy', label: 'Move or Copy...', command: 'sheet.move-copy', disabled: locked },
				{ id: 'tab-color', label: 'Tab Color', command: 'sheet.tab-color' },
				{ id: 'hide', label: 'Hide', command: 'sheet.hide', disabled: locked || single },
				{
					id: 'unhide',
					label: 'Unhide...',
					command: 'sheet.unhide',
					disabled: locked || state.hiddenSheets === 0,
				},
			],
			[
				{
					id: 'protect',
					label: state.protected ? 'Unprotect Sheet' : 'Protect Sheet...',
					command: 'review.protect-sheet',
				},
			],
		],
		state.readOnly,
	);
}

/** Keeps a menu of the given size inside the viewport, preferring the requested corner. */
export function clampToViewport(
	x: number,
	y: number,
	size: { width: number; height: number },
	viewport: { width: number; height: number },
	margin = 4,
): { left: number; top: number } {
	const fit = (position: number, extent: number, limit: number) =>
		Math.max(margin, Math.min(position, limit - extent - margin));
	return { left: fit(x, size.width, viewport.width), top: fit(y, size.height, viewport.height) };
}
