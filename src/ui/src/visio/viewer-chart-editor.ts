import {
	applyChartDataSheetEditsToSpace,
	applyChartEditIntent,
	chartSpaceDataSheet,
	renderChartSpaceSvg,
	type ChartSpace,
} from 'ooxml-core/chart';
import type { DrawingColorTheme } from 'ooxml-core/drawingml';
import { readEmbeddedChartSpace, renderChartSvg, writeEmbeddedChartSpace } from 'ooxml-core/xlsx';
import { readEmbeddedChart } from 'ooxml-core/xlsx/ui';
import { encodeBitmapEmf } from 'ooxml-core/visio';
import {
	defineChartDataSheet,
	type OfficeChartDataChangeEvent,
	type OfficeUiChartDataSheet,
} from '../chart/data-sheet';
import {
	defineChartFormatPane,
	type ChartFormatEditEvent,
	type ChartFormatTargetEvent,
	type OfficeUiChartFormatPane,
} from '../chart/format-pane';
import { parseSvgPreview } from '../ribbon/safe-svg';
import {
	rasterizeSvg,
	type ChartObjectData,
	type ChartTarget,
	type RasterizeSvg,
} from './viewer-insert-chart';

type Dialog = HTMLElement & { open: boolean; show(): void; close(): void };
type Button = HTMLElement & { disabled: boolean };

/** Which view the editor opens: the chart's data sheet or its Format pane. */
export type ChartEditMode = 'data' | 'format';

const PREVIEW_WIDTH = 640;
const LIVE = { width: 320, height: 200 };
const HEADINGS: Record<ChartEditMode, string> = { data: 'Edit Data', format: 'Format Chart' };

/**
 * Edit Data and Format Chart for an embedded chart object (Insert > Chart): the shared
 * `office-ui-chart-data-sheet` and `office-ui-chart-format-pane` over the chart part of the
 * embedded workbook, read as a `ChartSpace`. Data edits go through the data sheet model
 * (`applyChartDataSheetEditsToSpace`), formatting through `applyChartEditIntent`; a live preview
 * draws the edited part. OK writes the workbook again (`writeEmbeddedChartSpace`: the cells and
 * references follow a data edit), draws its metafile preview as the canvas will show it and hands
 * both to `submit`, which replaces the object. Formatting the part does not model (effects, 3-D,
 * text) is kept as it was.
 */
export class ChartObjectEditor {
	readonly dialog: Dialog;
	readonly sheet: OfficeUiChartDataSheet;
	readonly pane: OfficeUiChartFormatPane;
	readonly preview: HTMLElement;
	readonly error: HTMLElement;
	readonly #submit: (data: ChartObjectData, target: ChartTarget) => Promise<void>;
	readonly #rasterize: RasterizeSvg;
	readonly #buttons = new Map<string, Button>();
	#target: ChartTarget | undefined;
	#bytes: Uint8Array | undefined;
	#space: ChartSpace | undefined;
	#theme: DrawingColorTheme | undefined;
	#busy = false;
	constructor(
		root: ShadowRoot,
		submit: (data: ChartObjectData, target: ChartTarget) => Promise<void>,
		rasterize: RasterizeSvg = rasterizeSvg,
	) {
		this.#submit = submit;
		this.#rasterize = rasterize;
		const doc = root.ownerDocument;
		const registry = doc.defaultView?.customElements;
		defineChartDataSheet(registry);
		defineChartFormatPane(registry);
		this.dialog = doc.createElement('office-ui-dialog') as Dialog;
		this.dialog.className = 'chart-dialog chart-edit-dialog';
		this.sheet = doc.createElement('office-ui-chart-data-sheet') as OfficeUiChartDataSheet;
		this.pane = doc.createElement('office-ui-chart-format-pane') as OfficeUiChartFormatPane;
		this.preview = doc.createElement('div');
		this.preview.className = 'chart-edit-preview';
		this.preview.setAttribute('role', 'img');
		this.preview.setAttribute('aria-label', 'Chart preview');
		const body = doc.createElement('div');
		body.className = 'chart-edit-body';
		body.append(this.sheet, this.pane, this.preview);
		this.error = doc.createElement('p');
		this.error.setAttribute('role', 'alert');
		this.dialog.append(body, this.error);
		for (const name of ['Cancel', 'OK']) {
			const button = doc.createElement('office-ui-button') as Button;
			button.slot = 'footer';
			button.setAttribute('label', name);
			button.setAttribute('command', `chart-edit-${name.toLowerCase()}`);
			button.addEventListener('office-command', () => {
				if (button.disabled) return;
				if (name === 'Cancel') this.dialog.close();
				else void this.#save();
			});
			this.#buttons.set(name, button);
			this.dialog.append(button);
		}
		this.sheet.addEventListener('office-chart-data-change', (event) => {
			const { edits } = (event as OfficeChartDataChangeEvent).detail;
			if (this.#space) this.#update(applyChartDataSheetEditsToSpace(this.#space, edits));
		});
		this.pane.addEventListener('office-chart-format-target', (event) => {
			this.pane.target = (event as ChartFormatTargetEvent).detail.target;
		});
		this.pane.addEventListener('office-chart-format-edit', (event) => {
			const { intent } = (event as ChartFormatEditEvent).detail;
			if (!intent || !this.#space) return this.#update(this.#space);
			try {
				this.#update(applyChartEditIntent(this.#space, intent));
				this.error.textContent = '';
			} catch (error) {
				this.error.textContent = error instanceof Error ? error.message : String(error);
				this.#update(this.#space);
			}
		});
		this.pane.addEventListener('office-chart-format-close', () => this.dialog.close());
		root.append(this.dialog);
	}
	/** The edited chart part, while the editor is open. */
	get space(): ChartSpace | undefined {
		return this.#space;
	}
	/** The chart being edited while the editor is open, so a host can close it if the shape goes. */
	get target(): ChartTarget | undefined {
		return this.dialog.open ? this.#target : undefined;
	}
	/** Open `mode` for the chart object's workbook; false when the workbook holds no chart. */
	async open(mode: ChartEditMode, target: ChartTarget, workbook: Uint8Array): Promise<boolean> {
		const found = await readEmbeddedChartSpace(workbook);
		if (!found) return false;
		const scheme = found.scheme as unknown as Readonly<Record<string, string>>;
		this.#theme = { scheme: (name) => scheme[name] };
		this.#target = target;
		this.#bytes = workbook;
		this.pane.theme = this.#theme;
		this.pane.target = 'chartArea';
		this.sheet.hidden = mode !== 'data';
		this.pane.hidden = mode !== 'format';
		this.dialog.dataset.mode = mode;
		this.dialog.setAttribute('heading', HEADINGS[mode]);
		this.error.textContent = '';
		this.#update(found.space);
		this.dialog.show();
		return true;
	}
	close(): void {
		this.dialog.close();
	}
	#update(space: ChartSpace | undefined): void {
		this.#space = space;
		if (!space) return;
		this.sheet.data = chartSpaceDataSheet(space);
		this.pane.chartSpace = space;
		const svg = parseSvgPreview(this.dialog.ownerDocument, renderChartSpaceSvg(space, LIVE));
		if (svg) this.preview.replaceChildren(svg);
		else this.preview.replaceChildren();
	}
	async #save(): Promise<void> {
		const target = this.#target;
		const space = this.#space;
		const bytes = this.#bytes;
		if (this.#busy || !target || !space || !bytes) return;
		this.#busy = true;
		for (const button of this.#buttons.values()) button.disabled = true;
		try {
			const workbook = await writeEmbeddedChartSpace(bytes, space);
			const embedded = await readEmbeddedChart(workbook);
			if (!embedded) throw new Error('The chart could not be written.');
			const height = Math.min(
				1600,
				Math.max(120, Math.round((PREVIEW_WIDTH * target.height) / target.width)),
			);
			const raster = await this.#rasterize(
				this.dialog.ownerDocument,
				renderChartSvg(embedded.view, PREVIEW_WIDTH, height),
				PREVIEW_WIDTH,
				height,
			);
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
