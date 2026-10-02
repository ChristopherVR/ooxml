// Shared shapes of the Format Cells tabs.
import type { CellStyle, CellValue, DifferentialStyle, StylePatch } from '@christophervr/xlsx-core';
import type { Target } from '../../commands/util.js';
import type { EditorContext } from '../../context.js';

export type FormatTabId = 'number' | 'alignment' | 'font' | 'border' | 'fill' | 'protection';

export interface FormatCellsProps {
	tab?: FormatTabId;
	/** Initial Number category id (`accounting`, `date`, ...). */
	category?: string;
	/** Edit a conditional-format style instead of the selection; resolves with the new style. */
	dxf?: DifferentialStyle;
}

export interface TabInit {
	ctx: EditorContext;
	/** The resolved starting format (the active cell's, or the dxf as a cell style). */
	style: CellStyle;
	/** Editing a differential style (conditional formatting). */
	dxf: boolean;
	/** The active cell's value, for samples. */
	value: CellValue;
	date1904: boolean;
	props: FormatCellsProps;
	/** The selection, in cell mode. */
	target?: Target;
}

export interface FormatTab {
	id: FormatTabId;
	label: string;
	panel: HTMLElement;
	/** The user changed something on this tab. */
	dirty(): boolean;
	/** Style changes for `applyStyle` (cell mode), only the changed fields. */
	patch?(): StylePatch;
	/** Direct session edits (borders, merge) inside the dialog's batch, cell mode. */
	apply?(target: Target): void;
	/** Writes the tab's changes into a differential style (dxf mode). */
	toDxf?(dxf: DifferentialStyle): void;
	/** Focus target when the tab opens. */
	focus?(): void;
}
