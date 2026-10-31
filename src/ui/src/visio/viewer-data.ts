import { canEditDiagram } from './controller';
import type { VisioPage } from 'ooxml-core/visio';
import {
	VISIO_DATA_GRAPHIC_FILL_ROW,
	visioColorRules,
	visioAutoSpaceDataGraphicEdits,
	visioDataGraphicConfigs,
	visioDataGraphicEdits,
	visioDataGraphicFields,
	visioDataGraphicOwners,
	visioPositionDataGraphicEdits,
	VISIO_DATA_GRAPHIC_POSITIONS,
	type VisioDataGraphicPosition,
	visioLegendEdits,
	visioRemoveDataGraphicEdits,
	visioShapeDataRow,
	type VisioDataGraphicKind,
	visioSetDataGraphicItemEdits,
	type VisioDataGraphicItem,
} from 'ooxml-core/visio/ui';
import type { ViewerController, ViewerState } from './controller';
import type { VisioDataCommand } from './ribbon-action';
import type { RibbonCommand } from './ribbon-parts';
import { InsertDialog } from './viewer-insert-dialog';
import { ViewerDataSources } from './viewer-data-sources';
import { ViewerDataLinks } from './viewer-data-links';
import { ViewerShapeData } from './viewer-shape-data';
import { ViewerGraphicFields } from './viewer-data-fields';
import { ViewerAdvancedGraphics } from './viewer-data-advanced';
import { ViewerDataVisualizer } from './viewer-data-visualizer';
import {
	GRAPHIC_KINDS as KINDS,
	KIND_LABELS,
	POSITION_LABELS,
	kindLabel,
} from './viewer-data-labels';

type Run = (action: () => Promise<void>, success: string) => void;
const POSITION_PREFIX = 'graphic-position-';

/**
 * Visio's Data tab: Quick and Custom Import of a local workbook or CSV into a saved recordset,
 * Refresh All, the External Data window with row linking, data graphics, a legend and the
 * Define Shape Data dialog. Documents stay local; nothing is fetched.
 */
export class ViewerData {
	readonly sources: ViewerDataSources;
	readonly graphics: InsertDialog;
	readonly links: ViewerDataLinks;
	readonly shapeData: ViewerShapeData;
	readonly fields: ViewerGraphicFields;
	readonly advanced: ViewerAdvancedGraphics;
	readonly visualizer: ViewerDataVisualizer;
	/** Set while the item form edits one item of Advanced Data Graphics (undefined replacing: new). */
	#item: { replacing: VisioDataGraphicItem | undefined } | undefined;
	constructor(
		private readonly root: ShadowRoot,
		private readonly controller: ViewerController,
		private readonly announce: (message: string) => void,
		private readonly run: Run,
	) {
		this.graphics = new InsertDialog(
			root,
			'data-graphics-dialog',
			'Data Graphics',
			[
				{ name: 'field', label: 'Data field', choices: true },
				{ name: 'kind', label: 'Display as', choices: true },
				{ name: 'position', label: 'Position', choices: true },
				{ name: 'min', label: 'Minimum (blank for automatic)', input: 'number' },
				{ name: 'max', label: 'Maximum (blank for automatic)', input: 'number' },
			],
			['Apply', 'Cancel'],
			(button) => this.#applyGraphic(button),
		);
		this.graphics.choices('kind', Object.values(KINDS), 'text', {
			labels: KIND_LABELS,
			none: false,
		});
		this.graphics.choices('position', [...VISIO_DATA_GRAPHIC_POSITIONS], '', {
			labels: POSITION_LABELS,
			none: true,
			noneLabel: 'Default',
		});
		this.links = new ViewerDataLinks(root, controller, announce, run);
		this.sources = new ViewerDataSources(root, controller, announce, run, this.links);
		this.links.onRefresh = () => void this.sources.refresh();
		this.shapeData = new ViewerShapeData(root, controller, announce);
		this.fields = new ViewerGraphicFields(root, controller, run);
		this.visualizer = new ViewerDataVisualizer(controller, announce, run, this.sources.picker);
		this.advanced = new ViewerAdvancedGraphics(root, controller, announce, run, (page) =>
			this.#targets(page),
		);
		this.advanced.onEdit = (item) => this.#editItem(item);
	}
	/** Open the graphics form for one Advanced Data Graphics item, or a new one. */
	#editItem(item: VisioDataGraphicItem | undefined): void {
		const page = this.#page();
		if (!page) return;
		const fields = visioDataGraphicFields(page);
		this.graphics.choices(
			'field',
			fields.map((field) => field.name),
			item?.field ?? fields[0]!.name,
			{ labels: fields.map((field) => field.label), none: false },
		);
		(this.graphics.fields.get('kind') as HTMLSelectElement).value = item?.kind ?? 'text';
		(this.graphics.fields.get('position') as HTMLSelectElement).value = item?.position ?? '';
		(this.graphics.fields.get('min') as HTMLInputElement).value = String(item?.min ?? '');
		(this.graphics.fields.get('max') as HTMLInputElement).value = String(item?.max ?? '');
		this.graphics.error.textContent = '';
		this.#item = { replacing: item };
		this.advanced.dialog.dialog.close();
		this.graphics.dialog.show();
		this.graphics.fields.get('field')!.focus();
	}
	#editable(state: ViewerState): boolean {
		return canEditDiagram(state) && !state.loading && !state.edit.busy;
	}
	#page(): VisioPage | undefined {
		const state = this.controller.state;
		return state.document?.pages[state.pageIndex];
	}
	/** Selected top-level shapes, or every top-level shape on the page. */
	#targets(page: VisioPage): string[] {
		const selected = this.controller.state.selectedShapes
			.filter((shape) => !shape.pageId || shape.pageId === page.id)
			.map((shape) => shape.id)
			.filter((id) => page.shapes.some((shape) => shape.id === id));
		return selected.length ? selected : page.shapes.map((shape) => shape.id);
	}
	handle(command: VisioDataCommand): void {
		const state = this.controller.state;
		if (command === 'external-data-window') return this.links.toggle();
		if (command === 'graphic-fields') return this.fields.toggle();
		this.#item = undefined;
		if (!this.#editable(state)) return;
		if (command === 'define-shape-data') return this.shapeData.open();
		if (command === 'graphic-advanced') return this.advanced.open();
		if (command === 'visualizer-template') return this.visualizer.downloadTemplate();
		if (command === 'visualizer-refresh') return void this.visualizer.refresh();
		if (command.startsWith('visualizer-'))
			return void this.visualizer.create(
				command === 'visualizer-basic' ? 'basic' : 'cross-functional',
			);
		if (command === 'link-data') {
			this.links.show();
			if (!state.document?.dataRecordsets?.length) void this.sources.import(false);
			return;
		}
		if (command === 'quick-import' || command === 'custom-import')
			return void this.sources.import(command === 'custom-import');
		if (command === 'refresh') return void this.sources.refresh();
		const page = this.#page();
		if (!page) return;
		if (command === 'graphic-remove') {
			const targets = this.#targets(page);
			const all = targets.length === page.shapes.length;
			const edits = visioRemoveDataGraphicEdits(page, all ? [...targets, 'legend'] : targets);
			if (!edits.length) return this.announce('No data graphics to remove.');
			return this.run(() => this.controller.applyEdits(edits), 'Removed data graphics.');
		}
		if (
			command === 'graphic-auto-space' ||
			command === 'graphic-configure' ||
			command.startsWith(POSITION_PREFIX)
		)
			return this.#arrange(page, command);
		if (command === 'legend') return this.#legend(page, 'vertical');
		if (command === 'legend-horizontal') return this.#legend(page, 'horizontal');
		const fields = visioDataGraphicFields(page);
		if (!fields.length)
			return this.announce(
				'Link or define Shape Data first; data graphics display Shape Data fields.',
			);
		this.graphics.choices(
			'field',
			fields.map((field) => field.name),
			fields[0]!.name,
			{ labels: fields.map((field) => field.label), none: false },
		);
		(this.graphics.fields.get('kind') as HTMLSelectElement).value = KINDS[command] ?? 'text';
		for (const name of ['min', 'max'])
			(this.graphics.fields.get(name) as HTMLInputElement).value = '';
		(this.graphics.fields.get('position') as HTMLSelectElement).value = '';
		this.graphics.error.textContent = '';
		this.graphics.dialog.show();
		this.graphics.fields.get('field')!.focus();
	}
	/** Data Graphics Position, Configure and Auto Space act on graphics this viewer generated. */
	#arrange(page: VisioPage, command: VisioDataCommand): void {
		const targets = this.#targets(page);
		const owners = targets.filter((id) => visioDataGraphicConfigs(page, id).length);
		if (!owners.length)
			return this.announce('Select a shape with a text callout, data bar or icon set first.');
		if (command === 'graphic-configure') {
			const config = visioDataGraphicConfigs(page, owners[0]!)[0]!;
			this.graphics.choices(
				'field',
				visioDataGraphicFields(page).map((field) => field.name),
				config.field,
				{ labels: visioDataGraphicFields(page).map((field) => field.label), none: false },
			);
			(this.graphics.fields.get('kind') as HTMLSelectElement).value = config.kind;
			(this.graphics.fields.get('position') as HTMLSelectElement).value = config.position ?? '';
			(this.graphics.fields.get('min') as HTMLInputElement).value = String(config.min ?? '');
			(this.graphics.fields.get('max') as HTMLInputElement).value = String(config.max ?? '');
			this.graphics.error.textContent = '';
			this.graphics.dialog.show();
			this.graphics.fields.get('field')!.focus();
			return;
		}
		const edits =
			command === 'graphic-auto-space'
				? visioAutoSpaceDataGraphicEdits(page, owners)
				: visioPositionDataGraphicEdits(
						page,
						owners,
						command.slice(POSITION_PREFIX.length) as VisioDataGraphicPosition,
					);
		this.run(
			() => this.controller.applyEdits(edits),
			command === 'graphic-auto-space'
				? 'Spaced the data graphics so they do not overlap.'
				: 'Moved the data graphics.',
		);
	}
	#applyGraphic(button: string): void {
		if (button === 'Cancel') {
			this.graphics.dialog.close();
			if (this.#item) this.advanced.dialog.dialog.show();
			return;
		}
		const page = this.#page();
		if (!page) return;
		const number = (name: string) => {
			const text = this.graphics.value(name).trim();
			return text === '' ? undefined : Number(text);
		};
		const [min, max] = [number('min'), number('max')];
		if ([min, max].some((value) => value !== undefined && !Number.isFinite(value)))
			return void (this.graphics.error.textContent = 'Type numbers for the minimum and maximum.');
		const kind = this.graphics.value('kind') as VisioDataGraphicKind;
		const position = this.graphics.value('position') as VisioDataGraphicPosition | '';
		if (this.#item) return this.#applyItem(page, kind, position, min, max);
		const edits = visioDataGraphicEdits(page, this.#targets(page), {
			kind,
			...(position && kind !== 'color' ? { position } : {}),
			field: this.graphics.value('field'),
			...(min === undefined ? {} : { min }),
			...(max === undefined ? {} : { max }),
		});
		if (!edits.length)
			return void (this.graphics.error.textContent =
				'No shapes have a usable value for this field.');
		this.graphics.dialog.close();
		this.run(() => this.controller.applyEdits(edits), `Applied ${kindLabel(kind)} data graphics.`);
	}
	/** The item form of Advanced Data Graphics: add the item, or change the one being edited. */
	#applyItem(
		page: VisioPage,
		kind: VisioDataGraphicKind,
		position: VisioDataGraphicPosition | '',
		min: number | undefined,
		max: number | undefined,
	): void {
		const replacing = this.#item?.replacing;
		const edits = visioSetDataGraphicItemEdits(
			page,
			this.#targets(page),
			{
				kind,
				field: this.graphics.value('field'),
				...(position && kind !== 'color' ? { position } : {}),
				...(min === undefined ? {} : { min }),
				...(max === undefined ? {} : { max }),
			},
			replacing,
		);
		if (!edits.length)
			return void (this.graphics.error.textContent =
				'No shapes have a usable value for this field.');
		this.graphics.dialog.close();
		this.run(
			async () => {
				await this.controller.applyEdits(edits);
				this.advanced.refresh();
				this.advanced.dialog.dialog.show();
			},
			replacing ? 'Changed the data graphic item.' : 'Added the item to the data graphic.',
		);
	}
	#legend(page: VisioPage, orientation: 'vertical' | 'horizontal'): void {
		const coloured = page.shapes.filter((shape) =>
			visioShapeDataRow(shape, VISIO_DATA_GRAPHIC_FILL_ROW),
		);
		if (!coloured.length)
			return this.announce('Apply Color by Value first; the legend describes it.');
		const field = visioShapeDataRow(coloured[0]!, VISIO_DATA_GRAPHIC_FILL_ROW)!.label ?? '';
		const title = visioDataGraphicFields(page).find((item) => item.name === field)?.label ?? field;
		const rules = visioColorRules(
			page,
			coloured.map((shape) => shape.id),
			field,
		).rules;
		// One edit: an earlier legend is replaced, so the legend stays a single undo step.
		const edits = [
			...visioRemoveDataGraphicEdits(page, ['legend']),
			...visioLegendEdits(page, title, rules, orientation),
		];
		this.run(() => this.controller.applyEdits(edits), 'Inserted a legend for Color by Value.');
	}
	wire(viewport: HTMLElement): () => void {
		const disposeLinks = this.links.wire(viewport);
		const Abort = this.root.ownerDocument.defaultView?.AbortController ?? AbortController;
		const events = new Abort();
		this.root.addEventListener(
			'click',
			(event) => {
				if ((event.target as Element)?.closest?.('[data-define-shape-data]')) this.shapeData.open();
			},
			{ signal: events.signal },
		);
		return () => {
			events.abort();
			disposeLinks();
			for (const dialog of [
				this.sources.importer,
				this.graphics,
				this.advanced.dialog,
				this.shapeData.form,
			])
				dialog.dialog.close();
		};
	}
	#renderVisualizer(state: ViewerState, editing: boolean): void {
		const page = state.document?.pages[state.pageIndex];
		const set = (name: string, enabled: boolean, title?: string) => {
			const control = this.root.querySelector<RibbonCommand>(`[command="${name}"]`);
			if (!control) return;
			control.disabled = !enabled;
			if (title) control.title = title;
		};
		for (const id of ['visualizer-basic', 'visualizer-cross-functional'])
			set(id, editing && !!page);
		set(
			'visualizer-refresh',
			editing && this.visualizer.drawn(page),
			'Redraw the diagram from the edited table.',
		);
		const menu = this.root.querySelector<RibbonCommand>('[data-menu="data-visualizer"]');
		if (menu) menu.disabled = !editing || !page;
		set('advanced-data-graphics', editing && !!page && visioDataGraphicFields(page).length > 0);
	}
	render(state: ViewerState): void {
		const editing = this.#editable(state);
		const page = state.document?.pages[state.pageIndex];
		const button = (name: string) => this.root.querySelector<RibbonCommand>(`[command="${name}"]`);
		const set = (name: string, enabled: boolean, title?: string) => {
			const control = button(name);
			if (!control) return;
			control.disabled = !enabled;
			if (title) control.title = title;
		};
		const local = 'Open a .vsdx file to import data. Documents stay on this device.';
		set(
			'quick-import',
			editing && !!page,
			editing ? 'Import an Excel workbook or CSV file from this device.' : local,
		);
		set(
			'custom-import',
			editing && !!page,
			editing ? 'Choose the worksheet, range and header row to import.' : local,
		);
		const recordsets = state.document?.dataRecordsets?.length ?? 0;
		set(
			'refresh-all',
			editing && recordsets > 0,
			recordsets ? 'Read the source files again and update linked shapes.' : 'Import data first.',
		);
		const fields = page ? visioDataGraphicFields(page).length : 0;
		for (const id of [
			'graphic-text',
			'graphic-bar',
			'graphic-icon',
			'graphic-color',
			'graphic-remove',
		])
			set(id, editing && fields > 0);
		const arranged =
			editing &&
			!!page &&
			visioDataGraphicOwners(page).some((owner) => visioDataGraphicConfigs(page, owner).length > 0);
		for (const id of ['graphic-configure', 'graphic-auto-space', 'graphic-position'])
			set(id, arranged);
		for (const place of VISIO_DATA_GRAPHIC_POSITIONS) set(POSITION_PREFIX + place, arranged);
		const positionMenu = this.root.querySelector<RibbonCommand>('[data-menu="graphic-position"]');
		if (positionMenu) positionMenu.disabled = !arranged;
		const menuButton = this.root.querySelector<RibbonCommand>('[data-menu="data-graphics"]');
		if (menuButton) menuButton.disabled = !editing || !fields;
		const coloured = !!page?.shapes.some((shape) =>
			visioShapeDataRow(shape, VISIO_DATA_GRAPHIC_FILL_ROW),
		);
		set('legend-vertical', editing && coloured);
		set('legend-horizontal', editing && coloured);
		const legend = this.root.querySelector<RibbonCommand>('[data-menu="insert-legend"]');
		if (legend) {
			legend.disabled = !editing || !coloured;
			legend.title = coloured ? 'Insert Legend' : 'Insert Legend: apply Color by Value first.';
		}
		this.links.render(state);
		this.shapeData.render(state);
		this.fields.render(state);
		this.#renderVisualizer(state, editing);
		for (const dialog of [this.sources.importer, this.graphics, this.advanced.dialog])
			if (dialog.dialog.open && (!canEditDiagram(state) || state.loading)) dialog.dialog.close();
	}
}
