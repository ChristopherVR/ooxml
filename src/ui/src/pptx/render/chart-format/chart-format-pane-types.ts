/**
 * Framework-neutral descriptor of a chart Format task pane: icon tabs, then
 * collapsible sections of fields. The shared `pptx-ui-chart-format-pane`
 * element draws it; a field change goes through `applyChartFormatField`.
 *
 * @module render/chart-format/chart-format-pane-types
 */
import type { PptxChartFormatTargetKey } from 'ooxml-core/pptx';

import type { ChartFormatTargetOption, ChartFormatText } from './chart-format-targets';

/** The pane's icon tabs (PowerPoint: Fill & Line, Effects, Size & Properties, Options). */
export type ChartFormatTabId = 'fillLine' | 'effects' | 'options';

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
	| (FieldBase & { kind: 'text'; value: string });

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
}

/** The value a field change carries. */
export type ChartFormatFieldValue = string | number | boolean | '';
