// Context menus for the cell area, the row and column headers and the sheet tabs.
export {
	cellMenu,
	columnHeaderMenu,
	rowHeaderMenu,
	tabMenu,
	type CellMenuState,
	type HeaderMenuState,
	type MenuEntry,
	type TabMenuState,
} from 'ooxml-core/xlsx/ui';
export {
	currentContextMenu,
	openContextMenu,
	type ContextMenuHandle,
	type OpenMenuOptions,
} from './menu.js';
