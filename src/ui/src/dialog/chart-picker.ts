import { html, nothing } from 'lit';
import { OfficeElement, controlStyles } from '../base';
import { definer } from '../registry';
import { defineGallery, type OfficeGalleryPickEvent } from '../ribbon/gallery';
import { parseSvgPreview } from '../ribbon/safe-svg';
import css from './chart-picker.css?raw';

export interface OfficeChartType {
	id: string;
	label: string;
}

/** What the picker edits. `data` is only present when the host asked for the data editor. */
export interface OfficeChartValue {
	type: string;
	grouping: string;
	title: string;
	data?: string;
}

/**
 * Draws a chart as SVG markup for the current choices; `size` is the box in pixels. `tile` is true
 * for a chart type tile, which the gallery scales down: leave out the title and legend.
 */
export type OfficeChartPreview = (
	value: OfficeChartValue,
	size: { width: number; height: number },
	tile: boolean,
) => string;

export interface OfficeChartPickerLabels {
	chartType: string;
	grouping: string;
	title: string;
	data: string;
	dataHint: string;
	preview: string;
}

const ENGLISH: OfficeChartPickerLabels = {
	chartType: 'Chart type',
	grouping: 'Grouping',
	title: 'Chart title',
	data: 'Chart data',
	dataHint: 'One row per line, values separated by tabs or commas. The first row names the series.',
	preview: 'Preview',
};

const TILE = { width: 96, height: 60 };
const PREVIEW = { width: 360, height: 220 };

/**
 * `<office-ui-chart-picker>`: the body of every product's Insert Chart and Change Chart Type
 * dialog. It owns the chart type gallery, grouping, title, an optional data editor and a live
 * preview; the product owns the dialog around it, the OK and Cancel buttons and the chart
 * model. The host sets `types`, `groupings`, `groupedTypes` (the type ids that offer a grouping),
 * `labels` (translated) and `drawPreview`, which returns SVG for the current choices and is also
 * called once per type tile with `value.type` set to that tile. Read and write the choices with
 * `value`; leave `value.data` undefined when the numbers come from elsewhere (a spreadsheet
 * selection) and set it to a string to show the data editor. `change-only` hides the title and
 * data fields, as Change Chart Type keeps the chart's own.
 *
 * Emits `office-chart-change` with the value as `detail` on every edit.
 */
export class OfficeUiChartPicker extends OfficeElement {
	static override styles = controlStyles(css);
	static override properties = {
		changeOnly: { type: Boolean, attribute: 'change-only' },
	};
	/** Hides the title and data fields. */
	declare changeOnly: boolean;

	private choice: OfficeChartValue = { type: '', grouping: '', title: '' };
	private typeList: readonly OfficeChartType[] = [];
	private groupingList: readonly OfficeChartType[] = [];
	private grouped: readonly string[] = [];
	private text: OfficeChartPickerLabels = ENGLISH;
	private draw: OfficeChartPreview | undefined;

	constructor() {
		super();
		this.changeOnly = false;
	}

	get types(): readonly OfficeChartType[] {
		return this.typeList;
	}
	set types(value: readonly OfficeChartType[]) {
		this.typeList = value;
		if (!this.choice.type && value[0]) this.choice = { ...this.choice, type: value[0].id };
		this.requestUpdate('types');
	}
	get groupings(): readonly OfficeChartType[] {
		return this.groupingList;
	}
	set groupings(value: readonly OfficeChartType[]) {
		this.groupingList = value;
		if (!this.choice.grouping && value[0]) this.choice = { ...this.choice, grouping: value[0].id };
		this.requestUpdate('groupings');
	}
	get groupedTypes(): readonly string[] {
		return this.grouped;
	}
	set groupedTypes(value: readonly string[]) {
		this.grouped = value;
		this.requestUpdate('groupedTypes');
	}
	get labels(): OfficeChartPickerLabels {
		return this.text;
	}
	set labels(value: Partial<OfficeChartPickerLabels>) {
		this.text = { ...ENGLISH, ...value };
		this.requestUpdate('labels');
	}
	get drawPreview(): OfficeChartPreview | undefined {
		return this.draw;
	}
	set drawPreview(value: OfficeChartPreview | undefined) {
		this.draw = value;
		this.requestUpdate('drawPreview');
	}
	get value(): OfficeChartValue {
		return { ...this.choice };
	}
	set value(value: OfficeChartValue) {
		this.choice = { ...value };
		this.requestUpdate('value');
	}

	private edit(patch: Partial<OfficeChartValue>): void {
		this.choice = { ...this.choice, ...patch };
		this.requestUpdate('value');
		this.fire('office-chart-change', this.value);
	}
	private picked(event: Event): void {
		const id = (event as OfficeGalleryPickEvent).detail.itemId;
		if (this.typeList.some((type) => type.id === id)) this.edit({ type: id });
	}
	private svg(value: OfficeChartValue, tile: boolean): string {
		try {
			// A tile is drawn at the preview size and scaled by the gallery, so its text stays legible.
			return this.draw?.(value, PREVIEW, tile) ?? '';
		} catch {
			return '';
		}
	}

	protected override updated(): void {
		const gallery = this.renderRoot.querySelector<HTMLElement & { state: unknown }>(
			'office-ui-gallery',
		);
		if (gallery)
			gallery.state = {
				id: 'chart-type',
				label: this.text.chartType,
				sections: [
					{
						columns: 4,
						tileWidth: TILE.width,
						tileHeight: TILE.height,
						items: this.typeList.map((type) => ({
							id: type.id,
							label: type.label,
							applied: type.id === this.choice.type,
							preview: this.svg({ ...this.choice, type: type.id, title: '' }, true),
						})),
					},
				],
			};
		const preview = this.renderRoot.querySelector('.preview');
		if (preview) {
			const node = parseSvgPreview(this.ownerDocument, this.svg(this.choice, false));
			preview.replaceChildren(...(node ? [node] : []));
		}
	}

	protected override render() {
		defineGallery(this.ownerDocument.defaultView?.customElements);
		const text = this.text;
		const hasGrouping = this.groupingList.length > 0 && this.grouped.includes(this.choice.type);
		return html`<office-ui-gallery
				mode="panel"
				@office-gallery-pick=${(event: Event) => this.picked(event)}
			></office-ui-gallery
			>${
				hasGrouping
					? html`<label class="field"
							><span>${text.grouping}</span
							><select
								data-key="grouping"
								aria-label=${text.grouping}
								@change=${(event: Event) =>
									this.edit({ grouping: (event.target as HTMLSelectElement).value })}
								>${this.groupingList.map(
									(grouping) =>
										html`<option
											value=${grouping.id}
											?selected=${grouping.id === this.choice.grouping}
											>${grouping.label}</option
										>`,
								)}</select
							></label
						>`
					: nothing
			}${
				this.changeOnly
					? nothing
					: html`<label class="field"
							><span>${text.title}</span
							><input
								data-key="title"
								type="text"
								aria-label=${text.title}
								.value=${this.choice.title}
								@input=${(event: Event) => this.edit({ title: (event.target as HTMLInputElement).value })}
						/></label>`
			}${
				this.choice.data === undefined || this.changeOnly
					? nothing
					: html`<label class="field"
							><span>${text.data}</span
							><textarea
								data-key="data"
								rows="5"
								spellcheck="false"
								aria-label=${text.data}
								.value=${this.choice.data}
								@input=${(event: Event) =>
									this.edit({ data: (event.target as HTMLTextAreaElement).value })}
							></textarea
							><small>${text.dataHint}</small></label
						>`
			} <div class="preview" role="img" aria-label=${text.preview}></div>`;
	}
}

export const defineChartPicker = definer('office-ui-chart-picker', () => OfficeUiChartPicker);
