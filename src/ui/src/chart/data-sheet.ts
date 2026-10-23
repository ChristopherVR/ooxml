import { html, nothing } from 'lit';
import { live } from 'lit/directives/live.js';
import {
	applyChartDataSheetEdits,
	chartDataSheetCellEdit,
	chartDataSheetPasteEdits,
	type ChartDataSheet,
	type ChartDataSheetEdit,
} from 'ooxml-core/chart';
import { formatLocaleNumber, parseLocaleNumber, parseTsv } from 'ooxml-core/text';
import { OfficeElement, controlStyles } from '../base';
import { definer } from '../registry';
import { chartDataSheetKeyAction, type ChartDataSheetMode } from './data-sheet-keys';
import {
	CHART_DATA_SHEET_ENGLISH as ENGLISH,
	fill,
	type OfficeChartDataChange,
	type OfficeChartDataSheetHighlight,
	type OfficeChartDataSheetLabels,
} from './data-sheet-labels';
import css from './data-sheet.css?raw';

export type {
	OfficeChartDataChange,
	OfficeChartDataChangeEvent,
	OfficeChartDataSheetHighlight,
	OfficeChartDataSheetLabels,
} from './data-sheet-labels';

const EMPTY: ChartDataSheet = { categories: [], series: [] };

/**
 * `<office-ui-chart-data-sheet>`: the categories x series table of a chart's "Edit Data" view,
 * shared by every product. Series names run across the header row, category labels down the
 * first column, values in between; the toolbar adds a category or a series and each header has a
 * remove button while more than one is left. Arrow keys, Enter, Tab, F2 and Escape behave as in
 * the spreadsheet PowerPoint opens; pasting a tab-separated block (from Excel, Word or this sheet)
 * fills cells from the focused one, adding the series and categories it needs. Values are read and
 * shown in `locale`'s number convention (the browser's when unset, as Office follows the system's); a value that is not a number is refused and the cell
 * restored. The host sets `data` (`{ categories, series: [{ name, values }] }`), `labels`
 * (translated), `highlight` and `disabled`, and maps `office-chart-data-change`
 * `{ edits, data }` onto its own chart model; the sheet shows the edited data at once and takes
 * whatever `data` the host sets next.
 */
export class OfficeUiChartDataSheet extends OfficeElement {
	static override styles = controlStyles(css);
	static override properties = {
		data: { attribute: false },
		highlight: { attribute: false },
		disabled: { type: Boolean, reflect: true },
		locale: { type: String },
	};
	declare data: ChartDataSheet;
	declare highlight: OfficeChartDataSheetHighlight | null;
	declare disabled: boolean;
	declare locale: string;

	private text: OfficeChartDataSheetLabels = ENGLISH;
	private mode: ChartDataSheetMode = 'ready';
	private invalid = false;
	private highlightKey: string | null = null;

	constructor() {
		super();
		this.data = EMPTY;
		this.highlight = null;
		this.disabled = false;
		this.locale = '';
	}

	get labels(): OfficeChartDataSheetLabels {
		return this.text;
	}
	set labels(value: Partial<OfficeChartDataSheetLabels>) {
		this.text = { ...ENGLISH, ...value };
		this.requestUpdate('labels');
	}

	/** The number convention: `locale`, else the browser's, as Office follows the system locale. */
	private get language(): string | undefined {
		return this.locale || globalThis.navigator?.language || undefined;
	}

	private parse = (raw: string): number | null => parseLocaleNumber(raw, this.language);

	/** The text a cell shows for the current data. */
	private cellText(row: number, col: number): string {
		if (row === 0) return this.data.series[col - 1]?.name ?? '';
		if (col === 0) return this.data.categories[row - 1] ?? '';
		const value = this.data.series[col - 1]?.values[row - 1];
		return typeof value === 'number' ? formatLocaleNumber(value, this.language) : '';
	}

	/**
	 * Applies `edits` to the shown data and emits them. Product subclasses override this to map
	 * the edits onto their own model as well.
	 */
	protected emitEdits(edits: ChartDataSheetEdit[]): void {
		if (edits.length === 0 || this.disabled) return;
		this.data = applyChartDataSheetEdits(this.data, edits);
		const detail: OfficeChartDataChange = { edits, data: this.data };
		this.fire('office-chart-data-change', detail);
	}

	private position(input: HTMLInputElement): { row: number; col: number } {
		return { row: Number(input.dataset.row), col: Number(input.dataset.col) };
	}

	private commit(input: HTMLInputElement): void {
		const cell = this.position(input);
		if (input.value === this.cellText(cell.row, cell.col)) return;
		const edit = chartDataSheetCellEdit(cell, input.value, this.parse);
		this.invalid = edit === null;
		if (edit) this.emitEdits([edit]);
		this.requestUpdate();
	}

	private focusCell(row: number, col: number): void {
		const cell = this.renderRoot.querySelector<HTMLInputElement>(
			`input[data-row="${row}"][data-col="${col}"]`,
		);
		cell?.focus();
		cell?.select();
	}

	private onKeyDown(event: KeyboardEvent): void {
		// The editor around the sheet reads Delete, arrows and shortcuts as canvas commands.
		event.stopPropagation();
		const input = event.target as HTMLInputElement;
		const { row, col } = this.position(input);
		const rows = this.data.categories.length + 1;
		const cols = this.data.series.length + 1;
		const action = chartDataSheetKeyAction(event, this.mode, row, col, rows, cols);
		if (action.type === 'native') return;
		if (action.type === 'typing') {
			this.mode = 'enter';
			return;
		}
		event.preventDefault();
		if (action.type === 'edit') {
			this.mode = 'edit';
			input.setSelectionRange(input.value.length, input.value.length);
		} else if (action.type === 'revert') {
			input.value = this.cellText(row, col);
			this.mode = 'ready';
			this.invalid = false;
			this.requestUpdate();
			input.select();
		} else {
			this.commit(input);
			this.focusCell(action.row, action.col);
		}
	}

	private onFocus(event: FocusEvent): void {
		this.mode = 'ready';
		(event.target as HTMLInputElement).select();
	}

	private onPaste(event: ClipboardEvent): void {
		event.stopPropagation();
		const text = event.clipboardData?.getData('text/plain') ?? '';
		if (!/[\t\n\r]/.test(text.replace(/\r?\n$/, ''))) return;
		event.preventDefault();
		const anchor = this.position(event.target as HTMLInputElement);
		this.emitEdits(chartDataSheetPasteEdits(this.data, anchor, parseTsv(text), this.parse));
	}

	protected override updated(): void {
		const { highlight } = this;
		const key = highlight ? `${highlight.series}:${highlight.point ?? ''}` : null;
		if (key !== null && key !== this.highlightKey) {
			this.renderRoot.querySelector('.highlight')?.scrollIntoView?.({
				block: 'nearest',
				inline: 'nearest',
			});
		}
		this.highlightKey = key;
	}

	private cell(row: number, col: number, label: string, marked: boolean) {
		return html`<input
			type="text"
			class=${marked ? 'cell highlight' : 'cell'}
			inputmode=${row > 0 && col > 0 ? 'decimal' : 'text'}
			data-row=${row}
			data-col=${col}
			aria-label=${label}
			?disabled=${this.disabled}
			.value=${live(this.cellText(row, col))}
			@focus=${this.onFocus}
			@keydown=${this.onKeyDown}
			@change=${(event: Event) => this.commit(event.target as HTMLInputElement)}
			@paste=${this.onPaste}
			@copy=${(event: Event) => event.stopPropagation()}
			@cut=${(event: Event) => event.stopPropagation()}
		/>`;
	}

	private removeButton(label: string, edit: ChartDataSheetEdit) {
		return html`<button
			type="button"
			class="remove"
			title=${label}
			aria-label=${label}
			@click=${() => this.emitEdits([edit])}
			>×</button
		>`;
	}

	protected override render() {
		const { categories, series } = this.data;
		const t = this.text;
		const edit = !this.disabled;
		const mark = this.highlight;
		return html`<div class="head"
				><span class="heading">${t.heading}</span>${
					edit
						? html`<span class="actions"
								><button
									type="button"
									data-action="add-category"
									title=${t.addCategory}
									aria-label=${t.addCategory}
									@click=${() => this.emitEdits([{ kind: 'add-category' }])}
									>+ ${t.addCategoryShort}</button
								><button
									type="button"
									data-action="add-series"
									title=${t.addSeries}
									aria-label=${t.addSeries}
									@click=${() => this.emitEdits([{ kind: 'add-series' }])}
									>+ ${t.addSeriesShort}</button
								></span
							>`
						: nothing
				}</div
			><div class="summary"
				>${fill(t.summary, { series: series.length, categories: categories.length })}</div
			><div class="scroll"
				><table role="grid" aria-label=${t.heading}
					><thead
						><tr
							><th class="corner" aria-label=${t.categories}></th>${series.map(
								(entry, index) =>
									html`<th
										><span class="header"
											>${this.cell(
												0,
												index + 1,
												fill(t.seriesName, { number: index + 1 }),
												mark?.series === index && mark.point === undefined,
											)}${
												edit && series.length > 1
													? this.removeButton(t.removeSeries, {
															kind: 'remove-series',
															series: index,
														})
													: nothing
											}</span
										></th
									>`,
							)}</tr
						></thead
					><tbody
						>${categories.map(
							(_, row) =>
								html`<tr
									><th class="category"
										><span class="header"
											>${this.cell(row + 1, 0, fill(t.categoryLabel, { number: row + 1 }), false)}${
												edit && categories.length > 1
													? this.removeButton(t.removeCategory, {
															kind: 'remove-category',
															category: row,
														})
													: nothing
											}</span
										></th
									>${series.map(
										(entry, col) =>
											html`<td
												>${this.cell(
													row + 1,
													col + 1,
													fill(t.valueCell, { name: entry.name, number: row + 1 }),
													mark?.series === col && mark.point === row,
												)}</td
											>`,
									)}</tr
								>`,
						)}</tbody
					></table
				></div
			><div class="status" role="status">${this.invalid ? t.invalidNumber : ''}</div>`;
	}
}

export const defineChartDataSheet = definer(
	'office-ui-chart-data-sheet',
	() => OfficeUiChartDataSheet,
);
