import { canEditDiagram } from './controller';
import {
	inspectVisioRasterImage,
	VISIO_METADATA_TEXT_LIMIT,
	VISIO_PICTURE_MAX_BYTES,
	type VisioShape,
} from 'ooxml-core/visio';
import {
	editErrorMessage,
	isEditCancellation,
	normalizeVisioHyperlinkAddress,
	visioChartObjectInsertCommand,
	visioFollowTarget,
	visioPictureInsertCommand,
	visioSelectionIsOnPage,
} from 'ooxml-core/visio/ui';
import type { ViewerController, ViewerState } from './controller';
import type { VisioInsertItem } from './ribbon-action';
import type { RibbonCommand } from './ribbon-parts';
import { selectedShape, safeExternalHref } from './shape-inspector';
import { InsertDialog } from './viewer-insert-dialog';
import { InsertChartDialog, type ChartObjectData, type ChartTarget } from './viewer-insert-chart';

type Run = (action: () => Promise<void>, success: string) => void;

/**
 * Insert-tab Pictures, Link and ScreenTip, plus Ctrl+click link following. Pictures are read
 * from a local file only; links open in a new tab with noopener and never navigate the app.
 */
export class ViewerInsert {
	readonly input: HTMLInputElement;
	readonly link: InsertDialog;
	readonly tip: InsertDialog;
	readonly chart: InsertChartDialog;
	#target: { pageId: string; shapeId: string; row?: string; generation: number } | undefined;
	constructor(
		private readonly root: ShadowRoot,
		private readonly controller: ViewerController,
		private readonly announce: (message: string) => void,
		private readonly run: Run,
	) {
		const doc = root.ownerDocument;
		this.input = doc.createElement('input');
		this.input.type = 'file';
		this.input.accept = 'image/png,image/jpeg,image/gif,.png,.jpg,.jpeg,.gif';
		this.input.hidden = true;
		this.input.dataset.insertPicture = '';
		this.input.addEventListener('change', () => {
			const file = this.input.files?.[0];
			this.input.value = '';
			if (file) void this.#picture(file);
		});
		root.append(this.input);
		this.link = new InsertDialog(
			root,
			'link-dialog',
			'Hyperlinks',
			[
				{ name: 'address', label: 'Address', placeholder: 'https://example.com' },
				{ name: 'subAddress', label: 'Sub-address (page)', choices: true },
				{ name: 'description', label: 'Description' },
			],
			['OK', 'Remove Link', 'Cancel'],
			(button) => void this.#applyLink(button),
		);
		this.chart = new InsertChartDialog(root, (data, target) => this.#saveChart(data, target));
		this.tip = new InsertDialog(
			root,
			'screen-tip-dialog',
			'Shape ScreenTip',
			[{ name: 'text', label: 'ScreenTip text', maxLength: VISIO_METADATA_TEXT_LIMIT }],
			['OK', 'Cancel'],
			(button) => void this.#applyTip(button),
		);
	}
	#editable(state: ViewerState): boolean {
		return canEditDiagram(state) && !state.loading && !state.edit.busy;
	}
	#shape(state: ViewerState): VisioShape | undefined {
		const page = state.document?.pages[state.pageIndex];
		if (
			!page ||
			state.selectedShapes.length !== 1 ||
			!state.selectedShape ||
			!visioSelectionIsOnPage(state.selectedShape, page.id)
		)
			return undefined;
		return selectedShape(state.document, state.selectedShape, state.pageIndex);
	}
	open(item: VisioInsertItem): void {
		const state = this.controller.state;
		if (!this.#editable(state)) return;
		if (item === 'picture') {
			if (state.document?.pages[state.pageIndex]) this.input.click();
			return;
		}
		if (item === 'chart') {
			if (state.document?.pages[state.pageIndex]) this.chart.open();
			return;
		}
		const shape = this.#shape(state);
		const page = state.document?.pages[state.pageIndex];
		if (!shape || !page) {
			this.announce(`Select one shape to add a ${item === 'link' ? 'link' : 'ScreenTip'}.`);
			return;
		}
		const existing = (shape.hyperlinks ?? []).find((link) => !link.invisible);
		this.#target = {
			pageId: page.id,
			shapeId: shape.id,
			generation: this.controller.documentGeneration,
			...(item === 'link' && existing ? { row: existing.name } : {}),
		};
		const dialog = item === 'link' ? this.link : this.tip;
		dialog.error.textContent = '';
		if (item === 'link') {
			(this.link.fields.get('address') as HTMLInputElement).value = existing?.address ?? '';
			(this.link.fields.get('description') as HTMLInputElement).value = existing?.description ?? '';
			this.link.choices(
				'subAddress',
				state.document!.pages.map((item) => item.name),
				existing?.subAddress ?? '',
			);
		} else (this.tip.fields.get('text') as HTMLInputElement).value = shape.screenTip ?? '';
		dialog.dialog.show();
		this.render(state);
		dialog.fields.values().next().value?.focus();
	}
	close(): void {
		this.link.dialog.close();
		this.tip.dialog.close();
		this.chart.close();
	}
	/** Open the editor for the embedded chart `shapeId` on this page; false if it is not one. */
	editChart(shapeId: string): boolean {
		const state = this.controller.state;
		const page = state.document?.pages[state.pageIndex];
		const shape = page && embeddedChartShape(page.shapes, shapeId);
		if (!page || !shape?.embeddedObject || !this.#editable(state)) return false;
		void this.chart.edit(
			{
				pageId: page.id,
				shapeId,
				width: shape.width,
				height: shape.height,
				generation: this.controller.documentGeneration,
			},
			shape.embeddedObject.bytes,
		);
		return true;
	}
	/**
	 * Embed the dialog's workbook and preview: a new chart object on the page, or the existing one
	 * replaced. Rejects (so the dialog can say why) when the editor refuses.
	 */
	async #saveChart(data: ChartObjectData, target: ChartTarget | undefined): Promise<void> {
		const state = this.controller.state;
		const page = state.document?.pages[state.pageIndex];
		if (!page || !this.#editable(state)) throw new Error('Open a .vsdx file to insert a chart.');
		if (target) {
			if (target.pageId !== page.id || target.generation !== this.controller.documentGeneration)
				throw new Error('The drawing changed; open the chart again to edit it.');
			await this.#edit(
				() =>
					this.controller.applyEdits([
						{
							type: 'replace-chart-object',
							pageId: target.pageId,
							shapeId: target.shapeId,
							...data,
						},
					]),
				'Updated the chart.',
			);
			return;
		}
		const token = this.controller.captureCreationToken(page.id);
		const command = visioChartObjectInsertCommand(page, data.workbook, data.preview);
		await this.#edit(() => this.controller.applyCreationEdits([command], token), 'Inserted chart.');
	}
	/** Run an edit through the viewer's usual path, but resolve or reject with its outcome. */
	#edit(edit: () => Promise<void>, message: string): Promise<void> {
		return new Promise<void>((resolve, reject) =>
			this.run(async () => {
				try {
					await edit();
					resolve();
				} catch (error) {
					reject(error);
					throw error;
				}
			}, message),
		);
	}
	async #picture(file: File): Promise<void> {
		const state = this.controller.state;
		const page = state.document?.pages[state.pageIndex];
		if (!page || !this.#editable(state)) return;
		if (file.size > VISIO_PICTURE_MAX_BYTES) {
			this.announce(`${file.name} is larger than the 8 MB picture limit.`);
			return;
		}
		const token = this.controller.captureCreationToken(page.id);
		const bytes = new Uint8Array(await file.arrayBuffer());
		let command;
		try {
			command = visioPictureInsertCommand(page, inspectVisioRasterImage(bytes), bytes);
		} catch (error) {
			this.announce(
				`${file.name} cannot be inserted: ${error instanceof Error ? error.message : 'unreadable picture'} Use a PNG, JPEG or GIF file.`,
			);
			return;
		}
		this.run(
			() => this.controller.applyCreationEdits([command], token),
			`Inserted picture ${file.name}.`,
		);
	}
	async #apply(
		dialog: InsertDialog,
		edit: () => Parameters<ViewerController['applyEdits']>[0][number],
		message: string,
	): Promise<void> {
		const target = this.#target;
		if (!target || !dialog.dialog.open || this.controller.documentGeneration !== target.generation)
			return;
		try {
			await this.controller.applyEdits([edit()]);
			if (this.#target === target) dialog.dialog.close();
			this.announce(message);
		} catch (error) {
			if (dialog.dialog.open && this.#target === target && !isEditCancellation(error))
				dialog.error.textContent = editErrorMessage(error);
		}
	}
	async #applyLink(button: string): Promise<void> {
		if (button === 'Cancel') return this.link.dialog.close();
		const target = this.#target;
		if (!target) return;
		if (button === 'Remove Link') {
			if (target.row === undefined) return this.link.dialog.close();
			return this.#apply(
				this.link,
				() => ({ type: 'set-shape-hyperlink', ...target, row: target.row!, hyperlink: null }),
				'Removed the link.',
			);
		}
		const hyperlink = {
			address: normalizeVisioHyperlinkAddress(this.link.value('address')),
			subAddress: this.link.value('subAddress'),
			description: this.link.value('description'),
		};
		if (!hyperlink.address && !hyperlink.subAddress) {
			this.link.error.textContent = 'Type an address or choose a page.';
			return;
		}
		return this.#apply(
			this.link,
			() => ({
				type: 'set-shape-hyperlink',
				pageId: target.pageId,
				shapeId: target.shapeId,
				...(target.row === undefined ? {} : { row: target.row }),
				hyperlink,
			}),
			'Saved the link. Ctrl+click the shape to follow it.',
		);
	}
	async #applyTip(button: string): Promise<void> {
		if (button === 'Cancel') return this.tip.dialog.close();
		const target = this.#target;
		if (!target) return;
		const text = this.tip.value('text').replace(/[\r\n\t]+/g, ' ');
		return this.#apply(
			this.tip,
			() => ({ type: 'set-shape-screentip', pageId: target.pageId, shapeId: target.shapeId, text }),
			text ? 'Saved the ScreenTip.' : 'Removed the ScreenTip.',
		);
	}
	/** Ctrl+click follows a shape's link like Visio; Shift keeps additive selection. */
	wire(viewport: HTMLElement): () => void {
		const listener = (event: MouseEvent) => {
			if (!(event.ctrlKey || event.metaKey) || event.shiftKey || event.altKey) return;
			const group = (event.target as Element)?.closest?.<SVGGElement>('[data-shape-id]');
			const state = this.controller.state;
			if (!group || !state.document) return;
			const shape = selectedShape(
				state.document,
				{
					id: group.dataset.shapeId!,
					name: '',
					...(group.dataset.pageId ? { pageId: group.dataset.pageId } : {}),
				},
				state.pageIndex,
			);
			const target = shape && visioFollowTarget(state.document, shape);
			if (!target) return;
			event.preventDefault();
			event.stopPropagation();
			if (target.kind === 'page') return this.controller.setPage(target.index);
			const href = safeExternalHref(target.href);
			if (!href) return;
			viewport.ownerDocument.defaultView?.open(href, '_blank', 'noopener,noreferrer');
			this.announce(`Opened ${href} in a new tab.`);
		};
		viewport.addEventListener('click', listener, true);
		return () => {
			viewport.removeEventListener('click', listener, true);
			this.close();
		};
	}
	render(state: ViewerState): void {
		const editing = this.#editable(state);
		const page = state.document?.pages[state.pageIndex];
		const shape = editing ? this.#shape(state) : undefined;
		const button = (name: string) => this.root.querySelector<RibbonCommand>(`[command="${name}"]`);
		const pictures = button('pictures');
		if (pictures) {
			pictures.disabled = !editing || !page;
			pictures.title = editing
				? 'Insert a PNG, JPEG or GIF picture from this device (up to 8 MB).'
				: 'Open a .vsdx file to insert pictures.';
		}
		const chart = button('chart');
		if (chart) {
			chart.disabled = !editing || !page;
			chart.title = editing
				? 'Insert a chart drawn from data you type or paste. Double-click it to edit.'
				: 'Open a .vsdx file to insert a chart.';
		}
		for (const [name, label] of [
			['link', 'Link (Ctrl+K)'],
			['ctx-hyperlink', 'Hyperlink (Ctrl+K)'],
			['screen-tip', 'ScreenTip'],
		] as const) {
			const control = button(name);
			if (!control) continue;
			control.disabled = !shape;
			control.title = shape ? label : `${label}: select one shape on an editable drawing.`;
		}
		if (this.chart.dialog.open && (!canEditDiagram(state) || state.loading)) this.chart.close();
		for (const dialog of [this.link, this.tip]) {
			if (!dialog.dialog.open) continue;
			if (
				!canEditDiagram(state) ||
				state.loading ||
				this.controller.documentGeneration !== this.#target?.generation
			)
				dialog.dialog.close();
			else dialog.busy(state.edit.busy);
		}
		this.link.buttons.get('Remove Link')!.hidden = this.#target?.row === undefined;
	}
}

/** The shape `id` when it embeds a chart, looking inside groups. */
function embeddedChartShape(shapes: readonly VisioShape[], id: string): VisioShape | undefined {
	for (const shape of shapes) {
		if (shape.id === id) return shape.embeddedObject ? shape : undefined;
		const inner = embeddedChartShape(shape.children, id);
		if (inner) return inner;
	}
	return undefined;
}
