import {
	GROUPED_CHART_TYPES,
	INSERT_CHART_TYPES,
	gridChart,
	gridChartSvg,
	parseChartGrid,
} from 'ooxml-core/xlsx/ui';
import type { ChartType } from 'ooxml-core/xlsx';
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

/** Turns a chart's SVG into PNG bytes (Visio pictures are raster); replaceable for tests. */
export type RasterizeSvg = (
	doc: Document,
	svg: string,
	width: number,
	height: number,
) => Promise<Uint8Array>;

/** Size of the picture a chart becomes, in CSS pixels (drawn at twice this for sharpness). */
const CHART_SIZE = { width: 640, height: 400 };
const SCALE = 2;

const SAMPLE = 'Category\tSeries 1\tSeries 2\nA\t4\t2\nB\t3\t5\nC\t6\t4\nD\t2\t6';
const GROUPINGS = [
	{ id: 'clustered', label: 'Clustered' },
	{ id: 'stacked', label: 'Stacked' },
	{ id: 'percentStacked', label: '100% Stacked' },
];

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
		canvas.width = width * SCALE;
		canvas.height = height * SCALE;
		const context = canvas.getContext('2d');
		if (!context) throw new Error('This browser cannot draw the chart.');
		// A transparent PNG would show the page colour through the chart in dark themes.
		context.fillStyle = '#ffffff';
		context.fillRect(0, 0, canvas.width, canvas.height);
		context.drawImage(image, 0, 0, canvas.width, canvas.height);
		const blob = await new Promise<Blob | null>((resolve) => canvas.toBlob(resolve, 'image/png'));
		if (!blob) throw new Error('The chart could not be saved as a picture.');
		return new Uint8Array(await blob.arrayBuffer());
	} finally {
		win.URL.revokeObjectURL(url);
	}
};

/**
 * Insert > Chart: the shared chart picker (type, grouping, title and a data grid with a live
 * preview) in a dialog. OK draws the chart with the same core renderer Excel's charts use and
 * hands the PNG to `insert`, which places it on the page like any picture. The chart is a
 * picture: it is not an embedded workbook and cannot be edited as data afterwards.
 */
export class InsertChartDialog {
	readonly dialog: Dialog;
	readonly picker: Picker;
	readonly error: HTMLElement;
	readonly #insert: (png: Uint8Array) => Promise<void>;
	readonly #rasterize: RasterizeSvg;
	#busy = false;
	constructor(
		root: ShadowRoot,
		insert: (png: Uint8Array) => Promise<void>,
		rasterize: RasterizeSvg = rasterizeSvg,
	) {
		this.#insert = insert;
		this.#rasterize = rasterize;
		const doc = root.ownerDocument;
		defineChartPicker(doc.defaultView?.customElements);
		this.dialog = doc.createElement('office-ui-dialog') as Dialog;
		this.dialog.className = 'chart-dialog';
		this.dialog.setAttribute('heading', 'Insert Chart');
		this.picker = doc.createElement('office-ui-chart-picker') as Picker;
		this.picker.types = INSERT_CHART_TYPES.map(([id, label]) => ({ id, label }));
		this.picker.groupings = GROUPINGS;
		this.picker.groupedTypes = [...GROUPED_CHART_TYPES];
		this.picker.drawPreview = (value, size, tile) => {
			const grid = this.#grid(value);
			return grid ? gridChartSvg(grid, size.width, size.height, tile) : '';
		};
		this.error = doc.createElement('p');
		this.error.setAttribute('role', 'alert');
		this.dialog.append(this.picker, this.error);
		const buttons = new Map<string, Button>();
		for (const name of ['Cancel', 'Insert']) {
			const button = doc.createElement('office-ui-button') as Button;
			button.slot = 'footer';
			button.setAttribute('label', name);
			button.setAttribute('command', `chart-dialog-${name.toLowerCase()}`);
			button.addEventListener('office-command', () => {
				if (button.disabled) return;
				if (name === 'Cancel') this.dialog.close();
				else void this.#submit(buttons);
			});
			buttons.set(name, button);
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
	/** Open with the sample data, ready to be replaced. */
	open(): void {
		this.error.textContent = '';
		this.picker.value = { type: 'column', grouping: 'clustered', title: '', data: SAMPLE };
		this.dialog.show();
	}
	close(): void {
		this.dialog.close();
	}
	async #submit(buttons: ReadonlyMap<string, Button>): Promise<void> {
		if (this.#busy) return;
		const grid = this.#grid(this.picker.value);
		if (!grid || grid.chart.series.length === 0) {
			this.error.textContent =
				'Enter a header row and at least one row of numbers, separated by tabs or commas.';
			return;
		}
		this.#busy = true;
		for (const button of buttons.values()) button.disabled = true;
		try {
			const svg = gridChartSvg(grid, CHART_SIZE.width, CHART_SIZE.height);
			const png = await this.#rasterize(
				this.dialog.ownerDocument,
				svg,
				CHART_SIZE.width,
				CHART_SIZE.height,
			);
			await this.#insert(png);
			this.dialog.close();
		} catch (error) {
			this.error.textContent =
				error instanceof Error ? error.message : 'The chart was not inserted.';
		} finally {
			this.#busy = false;
			for (const button of buttons.values()) button.disabled = false;
		}
	}
}
