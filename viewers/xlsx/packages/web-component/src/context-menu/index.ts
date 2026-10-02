// Context menus for the cell area, the row and column headers and the sheet tabs.
export {
	cellMenu,
	clampToViewport,
	columnHeaderMenu,
	rowHeaderMenu,
	tabMenu,
	type CellMenuState,
	type HeaderMenuState,
	type MenuEntry,
	type TabMenuState,
} from './items.js';
export {
	currentContextMenu,
	openContextMenu,
	type ContextMenuHandle,
	type OpenMenuOptions,
} from './menu.js';
