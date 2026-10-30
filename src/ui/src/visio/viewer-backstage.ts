import type { ViewerState } from './controller';
import type { DocumentStorage, VisioFileCommand } from 'ooxml-core/embedding';
import { isVisioTemplateId, type VisioTemplateId, type VsdxSource } from 'ooxml-core/visio/ui';
import { backstageItems, type BackstagePage } from './backstage';
import { drawingExtension, openFromStorage, saveDrawingToStorage } from './viewer-file-commands';
import { wireNewDrawing } from './viewer-new-drawing';
import { createHostBackstagePages, type HostBackstagePages } from '../embedding/host-backstage';
import type { VisioRegionContext } from 'ooxml-core/visio/ui';
import { EMPTY_VISIO_CUSTOMIZATION, type ResolvedVisioCustomization } from './viewer-customization';
import {
	readPdfChoice,
	showPdfReport,
	syncPdfPanel,
	type PdfRangeChoice,
	type PdfReportSummary,
} from './viewer-pdf-options';

/** What the backstage needs from the element; every action delegates to existing APIs. */
export interface BackstageHost {
	root: ShadowRoot;
	viewport: HTMLElement;
	fileName(): string;
	load(source: VsdxSource): Promise<void>;
	createBlankDrawing(): Promise<void>;
	createFromTemplate(template: VisioTemplateId): Promise<void>;
	exportVsdx(): { bytes: Uint8Array; dirty: boolean };
	exportSvg(): { svg: string; pageIndex: number };
	/**
	 * A vector PDF of the chosen foreground pages with its per-page report, or a PNG of the
	 * current page (viewer-export).
	 */
	exportPicture(
		format: 'pdf' | 'png',
		pdf?: PdfRangeChoice,
	): Promise<{ blob: Blob; pageIndex: number; report?: PdfReportSummary }>;
	closeDocument(): void;
	showOptions(): void;
	announce(message: string): void;
	noteCount(): number;
	/** File > Print > Page Setup. */
	pageSetup?(): void;
	/** Announces the cancelable `file-command` event; false when the host handles the command. */
	fileCommand?(command: VisioFileCommand): boolean;
	/** Host-owned storage for Save and Open (the element's `storage`), when set. */
	storage?(): DocumentStorage | undefined;
	/** What the drawing is now; compared after an awaited save to see whether it changed. */
	revision?(): unknown;
	/** The drawing (as it was at `revision`) was saved to host storage under `fileName`. */
	savedToStorage?(fileName: string, revision: unknown): void;
	/** A host save failed (reported as `document-error`). */
	saveFailed?(error: Error): void;
	/** A copy was downloaded: the AutoRecover copy has done its job. */
	downloaded?(): void;
	/** The host customisation: removed pages and commands, and host pages. */
	customization?(): ResolvedVisioCustomization;
	/** The context of a host File tab page. */
	regionContext?(region: string): VisioRegionContext;
}

/**
 * Visio's File backstage: opened from the File tab, closed with Back or Escape (focus returns to
 * File). New, Open, Save, Save As, Export, Print and Close delegate to shared source APIs.
 */
type Backstage = HTMLElement & {
	items: unknown;
	open: boolean;
	show(page?: string): void;
	close(): void;
};
type Ribbon = HTMLElement & { focusFile(): void };

export class ViewerBackstage {
	readonly #root: Backstage;
	readonly #ribbon: Ribbon;
	readonly #input: HTMLInputElement;
	#state: ViewerState | undefined;
	#items = '';
	#urls = new Set<string>();
	#saving = false;
	#newDrawing: ReturnType<typeof wireNewDrawing> | undefined;
	readonly #hostPages: HostBackstagePages<VisioRegionContext>;
	constructor(private readonly host: BackstageHost) {
		this.#root = host.root.querySelector<Backstage>('office-ui-backstage')!;
		this.#hostPages = createHostBackstagePages<VisioRegionContext>(this.#root, {
			context: (region) => host.regionContext!(region),
			pageClass: 'backstage-page',
			label: '[visio-viewer]',
		});
		this.#ribbon = host.root.querySelector<Ribbon>('office-ui-ribbon')!;
		this.#input = this.#root.querySelector('.backstage-file-input')!;
	}
	get open(): boolean {
		return this.#root.open;
	}
	#custom(): ResolvedVisioCustomization {
		return this.host.customization?.() ?? EMPTY_VISIO_CUSTOMIZATION;
	}
	/** Re-applies the host customisation: hidden pages, host pages and their contexts. */
	refreshCustomization(): void {
		this.#hostPages.setPages(this.#custom().extraPages);
		this.#hostPages.refresh();
		if (this.#state) this.render(this.#state);
	}
	show(page: BackstagePage = 'info'): void {
		this.#ribbon.setAttribute('file-expanded', 'true');
		this.#root.show(page);
		if (page === 'print') this.#preview();
	}
	hide(): void {
		this.#root.close();
	}
	wire(): () => void {
		const Abort = this.host.root.ownerDocument.defaultView?.AbortController ?? AbortController;
		const events = new Abort();
		const options = { signal: events.signal };
		this.#newDrawing = wireNewDrawing(
			this.host.root,
			async (template) => {
				// Ctrl+N as well as the button: a host that removed New (`open: false`) gets neither.
				if (this.#custom().commands.has('new-blank')) return;
				if (!this.#announce('new')) return;
				if (template && template !== 'blank') await this.host.createFromTemplate(template);
				else await this.host.createBlankDrawing();
			},
			() => !!(this.#state?.loading || this.#state?.edit.busy),
			this.host.announce,
		);
		this.#ribbon.addEventListener(
			'office-ribbon-file',
			() => (this.open ? this.hide() : this.show()),
			options,
		);
		this.#root.addEventListener(
			'office-backstage-select',
			(event) => {
				const id = (event as CustomEvent<{ id: string }>).detail.id;
				if (this.#hostPages.show(id)) return;
				if (id === 'save') this.#download();
				else if (id === 'close') this.#close();
				else if (id === 'options') {
					// Visio opens its Options dialog over the drawing, not a backstage page.
					this.hide();
					this.host.showOptions();
				} else if (id === 'print') this.#preview();
			},
			options,
		);
		// Back, Escape or a command closed the backstage: return focus to File.
		this.#root.addEventListener(
			'office-backstage-close',
			() => {
				this.#ribbon.setAttribute('file-expanded', 'false');
				this.#ribbon.focusFile();
			},
			options,
		);
		this.#root.addEventListener(
			'click',
			(event) => {
				const action = (event.target as Element).closest?.<HTMLButtonElement>(
					'[data-backstage-action]',
				);
				if (action && !action.disabled) this.#run(action.dataset.backstageAction!);
			},
			options,
		);
		this.#input.addEventListener(
			'change',
			() => {
				const file = this.#input.files?.[0];
				this.#input.value = '';
				if (!file) return;
				this.hide();
				this.host
					.load(file)
					.catch((error: unknown) =>
						this.host.announce(error instanceof Error ? error.message : String(error)),
					);
			},
			options,
		);
		return () => {
			this.#newDrawing?.dispose();
			this.#hostPages.destroy();
			events.abort();
			for (const url of this.#urls) URL.revokeObjectURL(url);
			this.#urls.clear();
		};
	}
	#announce(command: VisioFileCommand): boolean {
		return this.host.fileCommand?.(command) ?? true;
	}
	#open(): void {
		if (!this.#announce('open')) return;
		const storage = this.host.storage?.();
		if (!storage?.open) return this.#input.click();
		openFromStorage(storage, this.host.root.ownerDocument)
			.then((file) => {
				if (!file) return;
				this.hide();
				return this.host.load(file);
			})
			.catch((error: unknown) =>
				this.host.announce(error instanceof Error ? error.message : String(error)),
			);
	}
	#run(action: string): void {
		if (this.#custom().commands.has(action)) return;
		if (action === 'open') this.#open();
		else if (action === 'new-blank') this.#newDrawing?.run();
		else if (action.startsWith('new-template:')) {
			const id = action.slice('new-template:'.length);
			if (isVisioTemplateId(id)) this.#newDrawing?.run(id);
		} else if (action === 'download') this.#download();
		else if (action === 'export-svg') this.#exportSvg();
		else if (action === 'export-pdf' || action === 'export-png')
			void this.#exportPicture(action === 'export-pdf' ? 'pdf' : 'png');
		else if (action === 'print') this.#print();
		else if (action === 'page-setup') {
			this.hide();
			this.host.pageSetup?.();
		}
	}
	#save(blob: Blob, name: string): void {
		const doc = this.host.root.ownerDocument;
		const url = URL.createObjectURL(blob);
		this.#urls.add(url);
		const anchor = doc.createElement('a');
		anchor.href = url;
		anchor.download = name;
		anchor.hidden = true;
		this.host.root.append(anchor);
		anchor.click();
		anchor.remove();
		// Keep the URL alive through browser download dispatch, then release it.
		doc.defaultView?.setTimeout(() => {
			if (this.#urls.delete(url)) URL.revokeObjectURL(url);
		}, 1000);
	}
	#base(): string {
		return this.host.fileName().replace(/\.vs[dt][xm]?$/i, '') || 'Drawing';
	}
	/** Save: download the edited drawing as a copy (File > Save and the Quick Access Toolbar). */
	download(): void {
		this.#download();
	}
	#download(): void {
		// One save at a time: a Save while a host save is pending is ignored, not queued. A host
		// that removed Save (`save: false`) gets neither `file-command` nor `storage.save`.
		if (
			!this.#state?.edit.sourceAvailable ||
			this.#saving ||
			this.#custom().commands.has('download')
		)
			return;
		if (!this.#announce('save')) return;
		const storage = this.host.storage?.();
		if (storage) return void this.#saveToStorage(storage);
		try {
			const result = this.host.exportVsdx();
			this.#save(
				new Blob([new Uint8Array(result.bytes)], { type: 'application/vnd.ms-visio.drawing' }),
				`${this.#base()}-${result.dirty ? 'edited' : 'original'}-copy.${drawingExtension(this.host.fileName())}`,
			);
			this.host.downloaded?.();
			this.hide();
			this.host.announce('Downloaded a copy of the drawing.');
		} catch (error) {
			this.host.announce(error instanceof Error ? error.message : String(error));
		}
	}
	/** True while a host save is pending. */
	get saving(): boolean {
		return this.#saving;
	}
	async #saveToStorage(storage: DocumentStorage): Promise<void> {
		this.#saving = true;
		try {
			const revision = this.host.revision?.();
			const { bytes } = this.host.exportVsdx();
			this.hide();
			const name = await saveDrawingToStorage(storage, bytes, this.host.fileName());
			this.host.savedToStorage?.(name, revision);
			this.host.announce('Saved the drawing.');
		} catch (cause) {
			const error = cause instanceof Error ? cause : new Error(String(cause));
			this.host.saveFailed?.(error);
			this.host.announce(error.message);
		} finally {
			this.#saving = false;
		}
	}
	#exportSvg(): void {
		try {
			const result = this.host.exportSvg();
			this.#save(
				new Blob([result.svg], { type: 'image/svg+xml;charset=utf-8' }),
				`${this.#base()}-page-${result.pageIndex + 1}.svg`,
			);
			this.hide();
			this.host.announce('SVG export requested. This is an approximate snapshot.');
		} catch (error) {
			this.host.announce(error instanceof Error ? error.message : String(error));
		}
	}
	async #exportPicture(format: 'pdf' | 'png'): Promise<void> {
		// A PDF keeps the Export page open: its options and the per-page report are on it.
		if (format === 'png') this.hide();
		else showPdfReport(this.#root, undefined);
		this.host.announce(`Preparing the ${format.toUpperCase()}...`);
		try {
			const count = this.#state?.document?.pages.filter((page) => !page.isBackground).length ?? 0;
			const choice = format === 'pdf' ? readPdfChoice(this.#root, count) : undefined;
			const { blob, pageIndex, report } = await this.host.exportPicture(format, choice);
			const name = format === 'pdf' ? this.#base() : `${this.#base()}-page-${pageIndex + 1}`;
			this.#save(blob, `${name}.${format}`);
			if (report) showPdfReport(this.#root, report);
			const degraded = report?.pages.filter(
				(page) => page.mode === 'picture' || page.degraded.length,
			);
			this.host.announce(
				format === 'png'
					? 'PNG export requested. Pages are pictures of the approximate rendering.'
					: degraded?.length
						? `PDF export requested. ${degraded.length} page(s) lost something; see the report on the Export page.`
						: 'PDF export requested.',
			);
		} catch (error) {
			this.host.announce(error instanceof Error ? error.message : String(error));
		}
	}
	/** The shared print preview shows an inert clone of the rendered page, never parsed markup. */
	#preview(): void {
		const preview = this.#root.querySelector<HTMLElement & { pages: Node[] }>(
			'office-ui-print-preview',
		)!;
		const paper = this.host.viewport.querySelector('svg.paper');
		preview.pages = paper ? [paper] : [];
	}
	#print(): void {
		const doc = this.host.root.ownerDocument;
		const paper = this.host.viewport.querySelector('svg.paper');
		if (!paper || !doc.body) return;
		const frame = doc.createElement('iframe');
		frame.setAttribute('aria-hidden', 'true');
		frame.style.cssText = 'position:fixed;width:0;height:0;border:0;visibility:hidden';
		doc.body.append(frame);
		const target = frame.contentDocument;
		const view = frame.contentWindow;
		if (!target || !view) return frame.remove();
		const clone = target.importNode(paper, true) as SVGSVGElement;
		clone.removeAttribute('style');
		clone.setAttribute('width', '100%');
		target.body.style.margin = '0';
		target.body.append(clone);
		view.addEventListener('afterprint', () => frame.remove(), { once: true });
		view.focus();
		view.print();
	}
	#close(): void {
		if (!this.#state?.document) return;
		this.hide();
		this.host.closeDocument();
	}
	render(state: ViewerState): void {
		this.#state = state;
		const page = state.document?.pages[state.pageIndex];
		const set = (key: string, value: string) => {
			const node = this.#root.querySelector(`[data-info="${key}"]`);
			if (node) node.textContent = value;
		};
		set('name', state.document ? this.host.fileName() || 'Untitled drawing' : 'No drawing open');
		set('pages', String(state.document?.pages.length ?? 0));
		syncPdfPanel(this.#root, state.document);
		set('page', page?.name ?? 'None');
		set('size', page ? `${page.width} × ${page.height} in` : 'None');
		set('shapes', page ? `${page.shapes.length} top-level` : 'None');
		set(
			'state',
			state.document?.format === 'vsd'
				? 'Legacy VSD preview (read only)'
				: !state.edit.sourceAvailable
					? state.document
						? 'Preview (read-only)'
						: 'None'
					: state.edit.dirty
						? 'Edited copy (not saved)'
						: 'Original',
		);
		const notes = this.host.noteCount();
		this.#root.querySelector('.backstage-note-count')!.textContent = notes
			? `${notes === 1 ? 'One part' : `${notes} parts`} of this drawing ${notes === 1 ? 'is' : 'are'} shown approximately or not at all:`
			: 'No compatibility notes.';
		const busy = state.loading || state.edit.busy;
		for (const tile of this.#root.querySelectorAll<HTMLButtonElement>(
			'[data-backstage-action="new-blank"], [data-backstage-action^="new-template:"]',
		)) {
			tile.disabled = busy;
			tile.title = busy ? 'Wait for the current document operation.' : '';
		}
		const saveReason =
			state.document?.format === 'vsd'
				? 'Legacy VSD drawings are read only here.'
				: 'Open a .vsdx file to save a copy.';
		const saveDisabled = !state.edit.sourceAvailable || busy;
		for (const download of this.#root.querySelectorAll<HTMLButtonElement>(
			'[data-backstage-action="download"]',
		)) {
			download.disabled = saveDisabled;
			download.title = saveDisabled ? saveReason : '';
		}
		const setup = this.#root.querySelector<HTMLButtonElement>(
			'[data-backstage-action="page-setup"]',
		);
		if (setup) {
			setup.disabled = !state.canEdit || !state.edit.sourceAvailable || busy || !page;
			setup.title = setup.disabled ? 'Open a .vsdx file to change the page setup.' : '';
		}
		// Rebuild the navigation only when Save or Close change, so focus stays put.
		const hidden = this.#custom().backstagePages;
		const items = this.#hostPages.items(
			backstageItems({
				save: { disabled: saveDisabled, ...(saveDisabled ? { title: saveReason } : {}) },
				close: { disabled: !state.document },
			}).filter((item) => !hidden.has(item.id)),
		);
		const key = JSON.stringify(items);
		if (key !== this.#items) {
			this.#items = key;
			this.#root.items = items;
		}
		for (const node of this.#root.querySelectorAll<HTMLButtonElement>(
			'[data-backstage-action^="export-"], [data-backstage-action="print"]',
		))
			node.disabled = !page || busy;
	}
}
