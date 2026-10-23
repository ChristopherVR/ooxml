/**
 * Framework-neutral descriptor of a chart Format task pane: icon tabs, then
 * collapsible sections of fields. `office-ui-chart-format-pane` draws it; the
 * product applies each field change (as a typed chart edit intent where it can).
 * Moved from PowerPoint's `render/chart-format/chart-format-pane-types.ts`.
 *
 * @module chart/format-pane-types
 */
import type { ChartFormatTarget, ChartFormatTargetKey } from 'ooxml-core/chart';

/** A translatable caption: key, `{{name}}` parameters and the English fallback. */
export interface ChartFormatText {
	key: string;
	fallback: string;
	params?: Readonly<Record<string, string | number>>;
	/** Parameters that are themselves captions ("{{axis}} Major Gridlines"). */
	paramTexts?: Readonly<Record<string, ChartFormatText>>;
}

/** One entry of the element picker. */
export interface ChartFormatTargetOption {
	key: ChartFormatTargetKey;
	target: ChartFormatTarget;
	label: ChartFormatText;
}

/**
 * The pane's icon tabs. Chart Options: Fill & Line, Effects and the element's
 * Options. Text Options: Text Fill & Outline, Text Effects and Textbox.
 */
export type ChartFormatTabId =
	| 'fillLine'
	| 'effects'
	| 'options'
	| 'textFill'
	| 'textEffects'
	| 'textBox';

/** Which half of the pane is showing: the element's own format or its text. */
export type ChartFormatPaneMode = 'chart' | 'text';

export interface ChartFormatTab {
	id: ChartFormatTabId;
	label: ChartFormatText;
}

/** One choice of a radio group or select. */
export interface ChartFormatChoice {
	value: string;
	label: ChartFormatText;
	disabled?: boolean;
	/** Why the choice is unavailable. */
	hint?: ChartFormatText;
}

interface FieldBase {
	/** Stable id, e.g. `fill.color`; what a field change carries. */
	id: string;
	label: ChartFormatText;
	disabled?: boolean;
	hint?: ChartFormatText;
}

export type ChartFormatField =
	| (FieldBase & { kind: 'radio'; value: string; choices: ChartFormatChoice[] })
	| (FieldBase & { kind: 'select'; value: string; choices: ChartFormatChoice[] })
	| (FieldBase & { kind: 'color'; value: string })
	| (FieldBase & {
			kind: 'number';
			/** `''` is Automatic for a field with `auto`. */
			value: number | '';
			min?: number;
			max?: number;
			step?: number;
			/** Unit suffix shown after the box ("pt", "%", "°"). */
			unit?: string;
			/** Empty means Automatic: the box shows the placeholder and clearing it is allowed. */
			auto?: boolean;
			/** Also draw a slider over `min`..`max` (Excel's spacing and transparency rows). */
			slider?: boolean;
	  })
	| (FieldBase & { kind: 'checkbox'; value: boolean })
	| (FieldBase & { kind: 'text'; value: string })
	| (FieldBase & {
			/** A file chooser; the change carries the chosen file as a `data:` URL. */
			kind: 'file';
			/** `accept` of the file input, e.g. `image/*`. */
			accept: string;
			/** Caption of the button that opens the chooser. */
			button: ChartFormatText;
			/** Whether a file is already chosen (the pane shows it as set). */
			value: boolean;
	  });

export interface ChartFormatSection {
	id: string;
	title: ChartFormatText;
	fields: ChartFormatField[];
	/** A note under the fields (an honest "not supported yet"). */
	note?: ChartFormatText;
	/** The host draws this section's body: the pane renders a `<slot>` of this name. */
	slot?: string;
}

export interface ChartFormatPaneModel {
	title: ChartFormatText;
	targetKey: ChartFormatTargetKey;
	target: ChartFormatTarget;
	targets: ChartFormatTargetOption[];
	/** Caption of the element picker. */
	picker: ChartFormatText;
	/** Empty when the source shows every section on one page. */
	tabs: ChartFormatTab[];
	activeTab: ChartFormatTabId;
	sections: ChartFormatSection[];
	/** Chart Options or Text Options. */
	mode: ChartFormatPaneMode;
	/** The element has text, so Text Options is available. */
	textOptions: boolean;
}

/** The value a field change carries. */
export type ChartFormatFieldValue = string | number | boolean | '';
