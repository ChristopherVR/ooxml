/**
 * Framework-neutral descriptor of a chart Format task pane: icon tabs, then
 * collapsible sections of fields. The shared `pptx-ui-chart-format-pane`
 * element draws it; a field change goes through `applyChartFormatField`.
 *
 * @module render/chart-format/chart-format-pane-types
 */
import type { PptxChartFormatTargetKey } from 'ooxml-core/pptx';

import type { ChartFormatTargetOption, ChartFormatText } from './chart-format-targets';

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
	/** Stable id, e.g. `fill.color`; what `applyChartFormatField` takes. */
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
}

export interface ChartFormatPaneModel {
	title: ChartFormatText;
	targetKey: PptxChartFormatTargetKey;
	targets: ChartFormatTargetOption[];
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
