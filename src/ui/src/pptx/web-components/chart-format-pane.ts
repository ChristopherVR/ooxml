import type { PptxChartData } from 'ooxml-core/pptx';

import {
	defineChartFormatPane,
	type OfficeUiChartFormatPane,
	type ChartFormatEditDetail,
} from '../../chart/format-pane';
import type { ChartFormatTranslate } from '../../chart/format-pane-common';
import { applyChartFormatEdit, pptxChartFormatKey, pptxChartFormatSource } from '../render';
import { CHART_FORMAT_PANE_STYLES } from './chart-format-pane-styles';
import { attachControlStyles } from './control-styles';
import { attachThemeEditorDock } from './theme-editor-dock';

export type ChartFormatPaneTranslate = (
	key: string,
	params?: Readonly<Record<string, string | number>>,
) => string;

export interface PptxUiChartFormatPaneElement extends HTMLElement {
	/** The selected chart's model. */
	chartData: PptxChartData | undefined;
	/** The element the pane formats (`chartArea`, `series:0`, `axis:1`, ...). */
	target: string;
	translateLabel: ChartFormatPaneTranslate;
	disabled: boolean;
}
/** A committed edit: the whole new chart model, for the host's update + history path. */
export type ChartFormatChangeEvent = CustomEvent<{ chartData: PptxChartData; target: string }>;
/** The element picker chose another element. */
export type ChartFormatTargetEvent = CustomEvent<{ target: string }>;
declare global {
	interface HTMLElementTagNameMap {
		'pptx-ui-chart-format-pane': PptxUiChartFormatPaneElement;
	}
}

const echo: ChartFormatPaneTranslate = (key) => key;

/**
 * `pptx-ui-chart-format-pane`: PowerPoint's chart Format task pane (Format
 * Chart Area, Format Data Series, Format Axis, ...), a thin subclass of the
 * shared `office-ui-chart-format-pane`. It reads `chartData` through
 * {@link pptxChartFormatSource} and maps each committed edit onto the pptx
 * chart operations (`applyChartFormatEdit`), firing `chart-format-change`
 * with the whole new chart model; the host applies it and keeps history.
 * Docked at the editor's right edge, a bottom sheet on narrow screens;
 * `inline` lays it out in flow. Keeps PowerPoint's tags, events and ids.
 */
export function definePptxChartFormatPane(registry: CustomElementRegistry): void {
	if (registry.get('pptx-ui-chart-format-pane')) {
		return;
	}
	defineChartFormatPane(registry);
	const Base = registry.get('office-ui-chart-format-pane') as typeof OfficeUiChartFormatPane;
	class ChartFormatPaneElement extends Base implements PptxUiChartFormatPaneElement {
		static override targetEvent = 'chart-format-target';
		static override closeEvent = 'chart-format-close';
		static override controlTags = {
			select: 'pptx-ui-select',
			radio: 'pptx-ui-radio',
			checkbox: 'pptx-ui-checkbox',
		};
		static override replayed = [...Base.replayed, 'chartData'];
		#data: PptxChartData | undefined;
		#t: ChartFormatPaneTranslate = echo;
		#undock?: () => void;
		constructor() {
			super();
			if (this.shadowRoot) {
				attachControlStyles(this.shadowRoot, CHART_FORMAT_PANE_STYLES);
			}
		}
		get chartData(): PptxChartData | undefined {
			return this.#data;
		}
		set chartData(value: PptxChartData | undefined) {
			this.#data = value;
			this.source = value && pptxChartFormatSource(value);
		}
		override get translateLabel(): ChartFormatPaneTranslate {
			return this.#t;
		}
		override set translateLabel(value: ChartFormatTranslate | undefined) {
			this.#t = value ?? echo;
			// Shared captions carry `chartFormat.*` keys; PowerPoint's dictionary has `pptx.chartFormat.*`.
			super.translateLabel = (key, params) => {
				const pptxKey = pptxChartFormatKey(key);
				const translated = this.#t(pptxKey, params);
				return translated === pptxKey ? key : translated;
			};
		}
		override connectedCallback(): void {
			super.connectedCallback();
			this.#undock = attachThemeEditorDock(this, { phoneSheet: true });
		}
		override disconnectedCallback(): void {
			super.disconnectedCallback();
			this.#undock?.();
		}
		protected override commit(edit: ChartFormatEditDetail): void {
			if (!this.#data) {
				return;
			}
			const next = applyChartFormatEdit(this.#data, edit);
			if (!next) {
				return;
			}
			this.chartData = next;
			this.dispatchEvent(
				new CustomEvent('chart-format-change', {
					detail: { chartData: next, target: edit.target },
					bubbles: true,
					composed: true,
				}),
			);
		}
	}
	registry.define('pptx-ui-chart-format-pane', ChartFormatPaneElement);
}
