import type { VisioPage } from 'ooxml-core/visio';
import {
	VISIO_VISUALIZER_COLUMNS,
	visioDataVisualizerEdits,
	visioParseVisualizerTable,
	visioVisualizerShapes,
	type VisioVisualizerVariant,
} from 'ooxml-core/visio/ui';
import { canEditDiagram, type ViewerController } from './controller';
import {
	readTable,
	readWorkbook,
	type DataFilePicker,
	type DataFileHandle,
} from './viewer-data-import';

type Run = (action: () => Promise<void>, success: string) => void;
const SAMPLE = [
	['1', 'Receive order', '2', '', 'Start', 'Sales'],
	['2', 'Check stock', '3,4', 'Yes,No', 'Decision', 'Warehouse'],
	['3', 'Ship order', '5', '', 'Process', 'Warehouse'],
	['4', 'Back order', '5', '', 'Process', 'Sales'],
	['5', 'Order complete', '', '', 'End', 'Sales'],
];
const csv = (row: readonly string[]) =>
	row.map((cell) => `"${cell.replace(/"/g, '""')}"`).join(',');

/**
 * Data Visualizer: draws a flowchart from a table of process steps (Process Step ID,
 * Description, Next Step ID, Connector Label, Shape Type, Function) read from a local CSV or
 * workbook, and redraws it from the edited table. Nothing leaves the device.
 */
export class ViewerDataVisualizer {
	#handle: DataFileHandle | undefined;
	constructor(
		private readonly controller: ViewerController,
		private readonly announce: (message: string) => void,
		private readonly run: Run,
		private readonly picker: DataFilePicker,
	) {}
	#page(): VisioPage | undefined {
		const state = this.controller.state;
		return state.document?.pages[state.pageIndex];
	}
	/** Whether this page already holds a flowchart this command drew. */
	drawn(page: VisioPage | undefined): boolean {
		return !!page && visioVisualizerShapes(page).steps.size > 0;
	}
	/** The kind of diagram a refresh redraws: the page's swimlanes decide. */
	#variantOf(page: VisioPage): VisioVisualizerVariant {
		return page.shapes.some(
			(shape) =>
				shape.structure?.type === 'container' && shape.structure.swimlane?.role === 'frame',
		)
			? 'cross-functional'
			: 'basic';
	}
	async create(variant: VisioVisualizerVariant): Promise<void> {
		const picked = await this.picker.pick();
		if (!picked) return;
		this.#handle = picked.handle;
		await this.#draw(picked.file, variant);
	}
	async refresh(): Promise<void> {
		const page = this.#page();
		if (!page || !this.drawn(page))
			return this.announce('Create a flowchart from a table first; Refresh Diagram redraws it.');
		let file: File | undefined;
		try {
			file = await this.#handle?.getFile();
		} catch {
			file = undefined;
		}
		if (!file) {
			this.announce(
				'Choose the edited table again; this browser cannot re-read the original file.',
			);
			const picked = await this.picker.pick();
			if (!picked) return;
			this.#handle = picked.handle;
			file = picked.file;
		}
		await this.#draw(file, this.#variantOf(page), true);
	}
	async #draw(file: File, variant: VisioVisualizerVariant, refreshing = false): Promise<void> {
		const page = this.#page();
		if (!page || !canEditDiagram(this.controller.state)) return;
		let plan;
		let warnings: string[];
		try {
			const table = readTable(await readWorkbook(file), { sheet: 0, range: '', header: true });
			const parsed = visioParseVisualizerTable(table);
			warnings = parsed.warnings;
			plan = visioDataVisualizerEdits(page, parsed.steps, variant);
			this.#count = parsed.steps.length;
		} catch (error) {
			return this.announce(
				`${file.name} cannot be drawn: ${error instanceof Error ? error.message : String(error)}`,
			);
		}
		const { batches, notes } = plan;
		const extra = [...warnings, ...notes];
		this.run(
			async () => {
				// One batch, so drawing or redrawing is one undo step.
				for (const batch of batches)
					try {
						await this.controller.applyEdits(batch);
					} catch (error) {
						// A page that stores no explicit size cannot be enlarged; the drawing still goes in.
						const sized = batch.filter((edit) => edit.type !== 'set-page-size');
						if (sized.length === batch.length) throw error;
						if (sized.length) await this.controller.applyEdits(sized);
					}
			},
			`${refreshing ? 'Redrew' : 'Drew'} ${this.#count} process steps from ${file.name}.${extra.length ? ` ${extra.join(' ')}` : ''}`,
		);
	}
	#count = 0;
	/** A CSV with the expected columns and a small example, for the person to edit. */
	downloadTemplate(): void {
		const doc = document;
		const text = [VISIO_VISUALIZER_COLUMNS, ...SAMPLE].map((row) => csv(row)).join('\r\n');
		const url = URL.createObjectURL(new Blob([text], { type: 'text/csv' }));
		const link = doc.createElement('a');
		link.href = url;
		link.download = 'data-visualizer-flowchart.csv';
		link.click();
		setTimeout(() => URL.revokeObjectURL(url), 1000);
		this.announce('Saved a table template with the Data Visualizer columns.');
	}
}
