import { defineChartPicker } from '../../dialog/chart-picker';
import type { OfficeChartValue, OfficeUiChartPicker } from '../../dialog/chart-picker';
import {
	DEFAULT_INSERT_CHART_KIND,
	GROUPED_INSERT_CHART_KINDS,
	INSERT_CHART_GROUPINGS,
	INSERT_CHART_TYPES,
	insertChartPreviewSvg,
} from '../render';
import type { InsertChartGrouping, InsertChartKind } from '../render';
import { createInsertDialogFrame, insertDialogText } from './insert-dialog-frame';
import type { InsertDialogFrame, InsertDialogTranslate } from './insert-dialog-frame';

/** What Insert hands the host: build the chart with `createDefaultChartElement(kind, _, rest)`. */
export interface InsertChartDialogDetail {
	kind: InsertChartKind;
	title: string;
	grouping: InsertChartGrouping;
}
export type InsertChartDialogInsertEvent = CustomEvent<InsertChartDialogDetail>;

export interface PptxUiInsertChartDialogElement extends HTMLElement {
	open: boolean;
	/** The chart type selected when the dialog opens. */
	kind: InsertChartKind;
	translateLabel: InsertDialogTranslate | undefined;
}

/** Sample title of a new chart, as `createDefaultChartElement` writes it. */
const DEFAULT_TITLE = 'Chart Title';

/**
 * `pptx-ui-insert-chart-dialog`: PowerPoint's Insert Chart dialog, the shared
 * `office-ui-chart-picker` (type gallery with live previews drawn by the slide's chart
 * renderer, grouping, title) in the shared dialog. Controlled: the host sets `open` (and
 * optionally `kind` and `translateLabel`); Insert emits `insert-chart` with
 * {@link InsertChartDialogDetail}, and every close (Insert, Cancel, Escape, the backdrop)
 * emits `dialog-close` so the host clears its open state. The host inserts the chart.
 */
export function definePptxInsertChartDialog(registry: CustomElementRegistry): void {
	if (registry.get('pptx-ui-insert-chart-dialog')) {
		return;
	}
	defineChartPicker(registry);

	class InsertChartDialog extends HTMLElement implements PptxUiInsertChartDialogElement {
		static observedAttributes = ['open'];
		#frame: InsertDialogFrame | undefined;
		#picker: OfficeUiChartPicker | undefined;
		#translate: InsertDialogTranslate | undefined;
		#kind: InsertChartKind = DEFAULT_INSERT_CHART_KIND;

		get open(): boolean {
			return this.hasAttribute('open');
		}
		set open(value: boolean) {
			this.toggleAttribute('open', Boolean(value));
		}
		get kind(): InsertChartKind {
			return this.#kind;
		}
		set kind(value: InsertChartKind) {
			this.#kind = INSERT_CHART_TYPES.some((entry) => entry.id === value)
				? value
				: DEFAULT_INSERT_CHART_KIND;
		}
		get translateLabel(): InsertDialogTranslate | undefined {
			return this.#translate;
		}
		set translateLabel(value: InsertDialogTranslate | undefined) {
			if (value === this.#translate) {
				return;
			}
			this.#translate = value;
			this.#sync();
		}

		connectedCallback(): void {
			this.#build();
			this.#sync();
			this.#show();
		}

		attributeChangedCallback(): void {
			if (this.isConnected) {
				this.#show();
			}
		}

		#text(key: string, fallback: string): string {
			return insertDialogText(this.#translate, key, fallback);
		}

		#build(): void {
			if (this.#frame) {
				return;
			}
			const picker = this.ownerDocument.createElement(
				'office-ui-chart-picker',
			) as OfficeUiChartPicker;
			picker.groupedTypes = [...GROUPED_INSERT_CHART_KINDS];
			picker.drawPreview = (value, size, tile) =>
				insertChartPreviewSvg(this.ownerDocument, value, size, tile, (key) => this.#text(key, key));
			this.#picker = picker;
			this.#frame = createInsertDialogFrame(
				this,
				picker,
				() => this.#insert(),
				() => this.#closed(),
			);
		}

		#sync(): void {
			const picker = this.#picker;
			if (!picker || !this.#frame) {
				return;
			}
			picker.types = INSERT_CHART_TYPES.map((entry) => ({
				id: entry.id,
				label: this.#text(entry.labelKey, entry.label),
			}));
			picker.groupings = INSERT_CHART_GROUPINGS.map((entry) => ({
				id: entry.id,
				label: this.#text(entry.labelKey, entry.label),
			}));
			picker.labels = {
				chartType: this.#text('pptx.ribbon.chartType', 'Chart type'),
				grouping: this.#text('pptx.chart.grouping', 'Grouping'),
				title: this.#text('pptx.chart.titlePlaceholder', 'Chart title'),
				preview: this.#text('pptx.insertChart.preview', 'Preview'),
			};
			this.#frame.sync(
				this.#text('pptx.insertChart.title', 'Insert Chart'),
				this.#text('pptx.smartart.close', 'Close'),
				this.#text('pptx.common.cancel', 'Cancel'),
				this.#text('pptx.common.insert', 'Insert'),
				true,
			);
		}

		#show(): void {
			const frame = this.#frame;
			const picker = this.#picker;
			if (!frame || !picker || frame.dialog.open === this.open) {
				return;
			}
			if (this.open) {
				const value: OfficeChartValue = {
					type: this.#kind,
					grouping: 'clustered',
					title: DEFAULT_TITLE,
				};
				picker.value = value;
			}
			frame.dialog.open = this.open;
			if (this.open) {
				picker.shadowRoot
					?.querySelector<HTMLElement>(`[data-gallery-item="${this.#kind}"]`)
					?.focus();
			}
		}

		#insert(): void {
			const value = this.#picker?.value;
			if (!value) {
				return;
			}
			const kind = INSERT_CHART_TYPES.find((entry) => entry.id === value.type)?.id;
			if (!kind) {
				return;
			}
			const grouping =
				INSERT_CHART_GROUPINGS.find((entry) => entry.id === value.grouping)?.id ?? 'clustered';
			this.dispatchEvent(
				new CustomEvent<InsertChartDialogDetail>('insert-chart', {
					detail: { kind, title: value.title, grouping },
					bubbles: true,
					composed: true,
				}),
			);
			if (this.#frame) {
				this.#frame.dialog.open = false;
			}
			this.#closed();
		}

		#closed(): void {
			this.removeAttribute('open');
			this.dispatchEvent(new CustomEvent('dialog-close', { bubbles: true, composed: true }));
		}
	}

	registry.define('pptx-ui-insert-chart-dialog', InsertChartDialog);
}
