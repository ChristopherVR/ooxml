import {
	GROUPED_CHART_TYPES,
	INSERT_CHART_TYPES,
	chartGridText,
	gridChart,
	gridChartSvg,
	gridChartWorkbook,
	parseChartGrid,
	readEmbeddedChart,
} from 'ooxml-core/xlsx/ui';
import type { ChartType } from 'ooxml-core/xlsx';
import { encodeBitmapEmf } from 'ooxml-core/visio';
import { defineChartPicker } from '../dialog/chart-picker';
import type { OfficeChartPreview, OfficeChartValue } from '../dialog/chart-picker';

type Dialog = HTMLElement & { open: boolean; show(): void; close(): void };
type Button = HTMLElement & { disabled: boolean };
type Picker = HTMLElement & {
	types: { id: string; label: string }[];
	groupings: { id: string; label: string }[];
	groupedTypes: string[];
	drawPreview: OfficeChartPreview;
	value: OfficeChartValue;
};

/** What the dialog hands back: the workbook to embed and its metafile preview. */
export interface ChartObjectData {
	workbook: Uint8Array;
	preview: Uint8Array;
}

/** The chart object being edited; absent for a new chart. */
export interface ChartTarget {
	pageId: string;
	shapeId: string;
	/** The shape's size in inches, so the preview keeps its proportions. */
	width: number;
	height: number;
	generation: number;
}

/** Draws an SVG document to pixels; replaceable for tests (a metafile preview needs raw pixels). */
export type RasterizeSvg = (
	doc: Document,
	svg: string,
	width: number,
	height: number,
) => Promise<{ width: number; height: number; rgba: Uint8ClampedArray }>;

/** Width of the preview in CSS pixels; the height follows the chart's proportions. */
const PREVIEW_WIDTH = 640;
const NEW_HEIGHT = 400;
const SCALE = 1.5;

const SAMPLE = 'Category\tSeries 1\tSeries 2\nA\t4\t2\nB\t3\t5\nC\t6\t4\nD\t2\t6';
const GROUPINGS = [
	{ id: 'clustered', label: 'Clustered' },
	{ id: 'stacked', label: 'Stacked' },
	{ id: 'percentStacked', label: '100% Stacked' },
];
const EDIT_NOTE =
	'Saving draws the chart again from the data above; formatting made in Excel is not kept.';

/** The browser's canvas draws the SVG; a chart is never loaded from a URL, only from this blob. */
export const rasterizeSvg: RasterizeSvg = async (doc, svg, width, height) => {
	const win = doc.defaultView;
	if (!win) throw new Error('There is no window to draw the chart in.');
	const url = win.URL.createObjectURL(new win.Blob([svg], { type: 'image/svg+xml' }));
	try {
		const image = new win.Image();
		image.src = url;
		await image.decode();
		const canvas = doc.createElement('canvas');
		canvas.width = Math.round(width * SCALE);
		canvas.height = Math.round(height * SCALE);
		const context = canvas.getContext('2d');
		if (!context) throw new Error('This browser cannot draw the chart.');
		// Transparent pixels would show the page colour through the preview in dark themes.
		context.fillStyle = '#ffffff';
		context.fillRect(0, 0, canvas.width, canvas.height);
		context.drawImage(image, 0, 0, canvas.width, canvas.height);
		const pixels = context.getImageData(0, 0, canvas.width, canvas.height);
		return { width: pixels.width, height: pixels.height, rgba: pixels.data };
	} finally {
		win.URL.revokeObjectURL(url);
	}
};

/**
 * Insert > Chart and Edit Chart: the shared chart picker (type, grouping, title and a data grid
 * with a live preview) in a dialog. OK writes an Excel workbook holding the chart, draws the same
 * chart as a metafile preview and hands both to `submit`, which embeds them as an object. The
 * chart is a real embedded workbook, so Visio itself can open it in Excel.
 */
export class InsertChartDialog {
	readonly dialog: Dialog;
	readonly picker: Picker;
	readonly error: HTMLElement;
	readonly #note: HTMLElement;
	readonly #submit: (data: ChartObjectData, target: ChartTarget | undefined) => Promise<void>;
	readonly #rasterize: RasterizeSvg;
	readonly #buttons = new Map<string, Button>();
	#target: ChartTarget | undefined;
	#busy = false;
	constructor(
		root: ShadowRoot,
		submit: (data: ChartObjectData, target: ChartTarget | undefined) => Promise<void>,
		rasterize: RasterizeSvg = rasterizeSvg,
	) {
		this.#submit = submit;
		this.#rasterize = rasterize;
		const doc = root.ownerDocument;
		defineChartPicker(doc.defaultView?.customElements);
		this.dialog = doc.createElement('office-ui-dialog') as Dialog;
		this.dialog.className = 'chart-dialog';
		this.picker = doc.createElement('office-ui-chart-picker') as Picker;
		this.picker.types = INSERT_CHART_TYPES.map(([id, label]) => ({ id, label }));
		this.picker.groupings = GROUPINGS;
		this.picker.groupedTypes = [...GROUPED_CHART_TYPES];
		this.picker.drawPreview = (value, size, tile) => {
			const grid = this.#grid(value);
			return grid ? gridChartSvg(grid, size.width, size.height, tile) : '';
		};
		this.#note = doc.createElement('p');
		this.#note.className = 'chart-note';
		this.error = doc.createElement('p');
		this.error.setAttribute('role', 'alert');
		this.dialog.append(this.picker, this.#note, this.error);
		for (const name of ['Cancel', 'Insert']) {
			const button = doc.createElement('office-ui-button') as Button;
			button.slot = 'footer';
			button.setAttribute('label', name);
			button.setAttribute('command', `chart-dialog-${name.toLowerCase()}`);
			button.addEventListener('office-command', () => {
				if (button.disabled) return;
				if (name === 'Cancel') this.dialog.close();
				else void this.#save();
			});
			this.#buttons.set(name, button);
			this.dialog.append(button);
		}
		root.append(this.dialog);
	}
	#grid(value: OfficeChartValue) {
		return gridChart(parseChartGrid(value.data ?? ''), value.type as ChartType, {
			grouping: value.grouping as 'clustered' | 'stacked' | 'percentStacked',
			...(value.title.trim() ? { title: value.title.trim() } : {}),
		});
	}
	#show(heading: string, action: string, note: string): void {
		this.error.textContent = '';
		this.#note.textContent = note;
		this.#note.hidden = !note;
		this.dialog.setAttribute('heading', heading);
		this.#buttons.get('Insert')!.setAttribute('label', action);
		this.dialog.show();
	}
	/** Open for a new chart, with sample data ready to be replaced. */
	open(): void {
		this.#target = undefined;
		this.picker.value = { type: 'column', grouping: 'clustered', title: '', data: SAMPLE };
		this.#show('Insert Chart', 'Insert', '');
	}
	/** Open an existing chart object with its own type, title and data. False if it has no chart. */
	async edit(target: ChartTarget, workbook: Uint8Array): Promise<boolean> {
		const embedded = await readEmbeddedChart(workbook);
		if (!embedded) return false;
		this.#target = target;
		this.picker.value = {
			type: embedded.chart.chartType,
			grouping: embedded.chart.grouping ?? 'clustered',
			title: embedded.chart.title ?? '',
			data: chartGridText(embedded.rows),
		};
		this.#show('Edit Chart', 'OK', EDIT_NOTE);
		return true;
	}
	close(): void {
		this.dialog.close();
	}
	/** The chart being edited while the dialog is open, so a host can close it if the shape goes. */
	get target(): ChartTarget | undefined {
		return this.dialog.open ? this.#target : undefined;
	}
	async #save(): Promise<void> {
		if (this.#busy) return;
		const value = this.picker.value;
		const grid = this.#grid(value);
		if (!grid || grid.chart.series.length === 0) {
			this.error.textContent =
				'Enter a header row and at least one row of numbers, separated by tabs or commas.';
			return;
		}
		this.#busy = true;
		for (const button of this.#buttons.values()) button.disabled = true;
		try {
			const target = this.#target;
			const height = target
				? Math.min(1600, Math.max(120, Math.round((PREVIEW_WIDTH * target.height) / target.width)))
				: NEW_HEIGHT;
			const raster = await this.#rasterize(
				this.dialog.ownerDocument,
				gridChartSvg(grid, PREVIEW_WIDTH, height),
				PREVIEW_WIDTH,
				height,
			);
			const workbook = await gridChartWorkbook(
				parseChartGrid(value.data ?? ''),
				value.type as ChartType,
				{
					grouping: value.grouping as 'clustered' | 'stacked' | 'percentStacked',
					...(value.title.trim() ? { title: value.title.trim() } : {}),
				},
			);
			if (!workbook) throw new Error('The chart could not be written.');
			await this.#submit(
				{ workbook, preview: encodeBitmapEmf(raster.width, raster.height, raster.rgba) },
				target,
			);
			this.dialog.close();
		} catch (error) {
			this.error.textContent = error instanceof Error ? error.message : 'The chart was not saved.';
		} finally {
			this.#busy = false;
			for (const button of this.#buttons.values()) button.disabled = false;
		}
	}
}
