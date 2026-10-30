/** ShapeSheet sections a shape shows or runs: control handles, right-click actions, data cells, tabs. */

/** One Controls row: a yellow handle in the shape's local y-up inches. */
export interface VisioControl {
	/** The row key the edit commands address: the row's IX, or its name when it has none. */
	row: string;
	/** The row name (`Row_1`), when it has one. */
	name?: string;
	x: number;
	y: number;
	/** The Prompt cell: the tip Visio shows while the handle is dragged. */
	prompt?: string;
	canGlue: boolean;
	/** Which axes a drag may change (XCon/YCon of 0 and a formula that holds a ratio or a constant). */
	drag: { x: boolean; y: boolean };
	/** Why neither axis can be dragged, when that is so. */
	locked?: string;
	/** True when the master supplies the row and the page does not override it. */
	inherited: boolean;
}

/** What running an action does, or why this editor does not run it. */
export type VisioActionRun =
	| { kind: 'setf'; targets: string[] }
	| { kind: 'unsupported'; reason: string };

/** One Actions row: a right-click menu entry of the shape. */
export interface VisioAction {
	/** The row key the edit commands address. */
	row: string;
	name?: string;
	/** The Menu text with the accelerator marker removed; the cached result of its formula. */
	menu: string;
	/** The Action cell's formula text, when it has one. */
	action?: string;
	checked: boolean;
	disabled: boolean;
	invisible: boolean;
	beginGroup: boolean;
	sortKey?: string;
	run: VisioActionRun;
}

/** One User-defined cell (`User.name`). */
export interface VisioUserCell {
	name: string;
	/** The cached value as saved. */
	value: string;
	formula?: string;
	prompt?: string;
}

/** One Scratch row; Scratch cells are named `X`, `Y` and `A` to `F`. */
export interface VisioScratchRow {
	row: string;
	values: Record<string, number>;
}

/** One tab stop of a paragraph, measured from the left edge of the text block in inches. */
export interface VisioTabStop {
	position: number;
	align: 'left' | 'center' | 'right' | 'decimal';
}

export interface VisioShapeSections {
	controls?: VisioControl[];
	actions?: VisioAction[];
	userCells?: VisioUserCell[];
	scratch?: VisioScratchRow[];
}
