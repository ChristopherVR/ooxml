import type { VisioDocument, VisioEdit } from 'ooxml-core/visio';
import type { VsdxSource } from './contract.js';
import { createWorkerParser } from './worker-parser.js';
import { ViewerController, type ViewerState } from './controller.js';
import { MAX_INPUT_BYTES } from './scene-validation.js';
import { selectedShape } from './shape-inspector.js';
import { wireViewerInputs } from './viewer-input.js';
import { renderLayerControls, wireLayerControls } from './viewer-layer-controls.js';
import { viewerStyles } from './styles.js';
import { canvasAndRibbonStyles } from './styles/index.js';
import { createRibbon } from './ribbon.js';
import { applyKeyTips } from './ribbon-keytips.js';
import { attachKeyTips } from 'ooxml-ui/controls';
import { createPageTabs, createStatusBar } from './status-bar.js';
import { createFindBar, wireFindBar } from './viewer-search.js';
import { fitZoom } from './viewer-fit.js';
import { createShapesStrip, createShapesWindow } from './shapes-window.js';
import { createBackstage, type BackstagePage } from './backstage.js';
import { ViewerBackstage } from './viewer-backstage.js';
import { createContextMenus, wireContextMenus } from './viewer-context-menu.js';
import { wireTellMe } from './viewer-tell-me.js';
import { createPanZoom, ViewerPanZoom } from './viewer-pan-zoom.js';
import { wireStencil } from './viewer-stencil.js';
import { createRulers, type Rulers } from './viewer-ruler.js';
import { ViewerEditControls } from './viewer-edit-controls.js';
import { searchControls, renderSearchControls, type SearchControls } from './viewer-search.js';
import { ViewerChrome, viewerChromeTemplate } from './viewer-chrome.js';
import { ViewerCanvas } from './viewer-canvas.js';
import { ViewerCommands } from './viewer-commands.js';
import { editErrorMessage } from './edit-error.js';
import { registerViewerControls } from './office-ui.js';
import { exportPageSvg, type SvgExportOptions, type SvgExportResult } from './export-svg.js';
import {
	createPrintSnapshot,
	type CurrentPagePrintSnapshotOptions,
	type PrintSnapshot,
} from './print-snapshot.js';

const BaseElement = (
	typeof HTMLElement === 'undefined' ? class {} : HTMLElement
) as typeof HTMLElement;
export class VisioViewerElement extends BaseElement {
	readonly controller = new ViewerController(createWorkerParser());
	#root: ShadowRoot;
	#disposeInputs: () => void;
	#viewport: HTMLDivElement;
	#zoomSlider: HTMLElement & { value: number; disabled: boolean };
	#shapeStatus: HTMLElement;
	#announcement: string | undefined;
	#commands: ViewerCommands;
	#rulers: Rulers;
	#canvas: ViewerCanvas;
	#panZoom: ViewerPanZoom;
	#backstage: ViewerBackstage;
	#fileName = '';
	#loadToken = 0;
	#findBar: HTMLElement;
	#status: HTMLSpanElement;
	#diagnostics: HTMLSpanElement;
	#toolbar: HTMLDivElement;
	#chrome: ViewerChrome;
	#search: SearchControls;
	#edit: ViewerEditControls;
	#layers: HTMLDetailsElement;
	#notes: HTMLUListElement;
	#unsubscribe: () => void;
	#eventUnsubscribe: () => void;
	#disposed = false;
	#suspended = false;
	#fontEvents: FontFaceSet | undefined;
	#fontsChanged = () => {
		if (this.#disposed) return;
		this.#canvas.invalidate();
		this.#render(this.controller.state);
	};
	constructor() {
		super();
		this.#root = this.attachShadow({ mode: 'open' });
		// Shared Office controls must be defined before the static template upgrades them.
		registerViewerControls();
		// This template is static, never document content.
		this.#root.innerHTML = `<style>${viewerStyles}${canvasAndRibbonStyles}</style>${viewerChromeTemplate}`;
		const workspace = this.#root.querySelector('.workspace')!;
		this.#findBar = createFindBar(document);
		const ribbon = createRibbon(document);
		applyKeyTips(ribbon);
		workspace.before(ribbon, this.#findBar);
		workspace.prepend(createShapesStrip(document), createShapesWindow(document));
		workspace.append(createPanZoom(document));
		this.#root.append(createBackstage(document), ...createContextMenus(document));
		workspace.after(createPageTabs(document), createStatusBar(document));
		this.#viewport = this.#root.querySelector('.viewport')!;
		this.#rulers = createRulers(this.#viewport);
		this.#zoomSlider = this.#root.querySelector('office-ui-zoom-slider')!;
		this.#shapeStatus = this.#root.querySelector('[data-shape-status]')!;
		this.#status = this.#root.querySelector('[data-status]')!;
		this.#diagnostics = this.#root.querySelector('[data-diagnostics]')!;
		this.#toolbar = this.#root.querySelector('.toolbar')!;
		this.#search = searchControls(this.#root);
		this.#edit = new ViewerEditControls(this.#root, this.controller);
		this.#layers = this.#root.querySelector('.layer-controls')!;
		this.#notes = this.#root.querySelector('.notes ul')!;
		this.#canvas = new ViewerCanvas(
			this.#viewport,
			this.#notes,
			this.#root.querySelector('.notes')!,
			this.#root.querySelector('.shape-inspector')!,
		);
		this.#chrome = new ViewerChrome(this.#root, this.controller);
		this.#panZoom = new ViewerPanZoom(this.#root, this.#viewport, this.controller, (open) =>
			this.#root.querySelector('[command="pan-zoom"]')?.setAttribute('checked', String(open)),
		);
		this.#commands = new ViewerCommands({
			root: this.#root,
			viewport: this.#viewport,
			controller: this.controller,
			fit: (mode) => this.#fit(mode),
			togglePane: (pane) => this.#chrome.togglePane(pane),
			reveal: (panel, focusText) => this.#chrome.reveal(panel, focusText),
			rulers: this.#rulers,
			togglePanZoom: () => this.#panZoom.toggle(),
			focusSearch: () => {
				this.#chrome.closeCompactTools();
				this.#findBar.hidden = false;
				this.#search.input.focus();
				this.#search.input.select();
			},
			announce: (message) => {
				this.#announcement = message;
				this.#status.textContent = message;
			},
		});
		this.#backstage = new ViewerBackstage({
			root: this.#root,
			viewport: this.#viewport,
			fileName: () => this.#fileName,
			load: (source) => this.load(source),
			exportVsdx: () => this.controller.exportVsdx(),
			exportSvg: () => this.exportSvg(),
			closeDocument: () => {
				this.#fileName = '';
				this.controller.setDocument(null);
			},
			revealNotes: () => this.#chrome.reveal('notes'),
			announce: (message) => {
				this.#announcement = message;
				this.#status.textContent = message;
			},
			noteCount: () => this.#notes.children.length,
		});
		this.#disposeInputs = this.#wireInputs();
		this.#unsubscribe = this.controller.subscribe((state) => this.#render(state));
		this.#eventUnsubscribe = this.controller.onEvent((name, detail) => {
			this.dispatchEvent(new CustomEvent(name, { detail, bubbles: true, composed: true }));
		});
	}
	get document(): VisioDocument | null {
		return this.controller.state.document;
	}
	set document(value: VisioDocument | null) {
		++this.#loadToken;
		this.#fileName = '';
		this.controller.setDocument(value);
	}
	get pageIndex(): number {
		return this.controller.state.pageIndex;
	}
	set pageIndex(value: number) {
		this.controller.setPage(value);
	}
	get zoom(): number {
		return this.controller.state.zoom;
	}
	set zoom(value: number) {
		this.controller.setZoom(value);
	}
	get showToolbar(): boolean {
		return !this.#toolbar.hidden;
	}
	set showToolbar(value: boolean) {
		this.#assertAlive();
		this.#toolbar.hidden = !value;
		this.#root.querySelector<HTMLElement>('.zoom-controls')!.hidden = !value;
	}
	get renderWarnings(): readonly string[] {
		return this.#canvas.warnings;
	}
	async load(source: VsdxSource): Promise<void> {
		this.#assertAlive();
		const token = ++this.#loadToken;
		if (source instanceof Blob) {
			await this.controller.loadSource(() => {
				if (source.size > MAX_INPUT_BYTES)
					throw new Error('This viewer accepts files up to 32 MiB.');
				return source.arrayBuffer();
			});
		} else await this.controller.load(source);
		// A destroyed or superseded viewer keeps no name from a late load.
		if (this.#disposed || token !== this.#loadToken) return;
		this.#fileName = typeof File !== 'undefined' && source instanceof File ? source.name : '';
		this.#render(this.controller.state);
	}
	/** Leave Visio's File backstage and return to the drawing. */
	closeBackstage(): void {
		this.#assertAlive();
		this.#backstage.hide();
	}
	/** Open Visio's File backstage at a page (Info by default). */
	openBackstage(page?: BackstagePage): void {
		this.#assertAlive();
		this.#backstage.show(page);
	}
	/** Name of the last opened local file, or empty for bytes, models and closed drawings. */
	get fileName(): string {
		return this.#fileName;
	}
	applyEdits(edits: readonly VisioEdit[]): Promise<void> {
		this.#assertAlive();
		return this.controller.applyEdits(edits);
	}
	replacePlainText(pageId: string, shapeId: string, text: string): Promise<void> {
		this.#assertAlive();
		return this.controller.replacePlainText(pageId, shapeId, text);
	}
	undo(): Promise<void> {
		this.#assertAlive();
		return this.controller.undo();
	}
	redo(): Promise<void> {
		this.#assertAlive();
		return this.controller.redo();
	}
	cancelEdit(): void {
		this.#assertAlive();
		this.#edit.reset(false);
	}
	exportVsdx(): ReturnType<ViewerController['exportVsdx']> {
		this.#assertAlive();
		return this.controller.exportVsdx();
	}
	fit(): void {
		this.#assertAlive();
		this.#fit('page');
	}
	/** Visio Fit to Window shows the whole page; Page Width fills the canvas width. */
	#fit(mode: 'page' | 'width'): void {
		const page = this.document?.pages[this.pageIndex];
		if (page) this.zoom = fitZoom(this.#viewport, page, mode);
	}
	setLayerVisibility(pageId: string, layerId: string, visible: boolean | null): void {
		this.#assertAlive();
		this.controller.setLayerVisibility(pageId, layerId, visible);
	}
	resetLayerVisibility(pageId?: string): void {
		this.#assertAlive();
		this.controller.resetLayerVisibility(pageId);
	}
	/** Return a portable saved-display current-page snapshot without changing selection or downloading a file. */
	exportSvg(options?: SvgExportOptions): SvgExportResult {
		this.#assertAlive();
		const { document, pageIndex } = this.controller.state;
		if (!document) throw new Error('Open a document before exporting SVG.');
		return exportPageSvg(document, pageIndex, options);
	}
	/** Prepare only the captured current drawing with saved display visibility. No frame, download, print dialog or state change. */
	createPrintSnapshot(options?: CurrentPagePrintSnapshotOptions): PrintSnapshot {
		this.#assertAlive();
		if (
			options !== undefined &&
			(!options ||
				typeof options !== 'object' ||
				Array.isArray(options) ||
				'pageIndices' in options)
		)
			throw new Error('Current-page print snapshot options may only contain limits.');
		const { document, pageIndex } = this.controller.state;
		const generation = this.controller.documentGeneration;
		if (!document) throw new Error('Open a document before preparing a print snapshot.');
		const page = document.pages[pageIndex];
		const result = createPrintSnapshot(document, { ...options, pageIndices: [pageIndex] });
		this.#assertAlive();
		if (generation !== this.controller.documentGeneration || document.pages[pageIndex] !== page)
			throw new Error('The document changed while preparing the print snapshot.');
		return result;
	}
	destroy(): void {
		if (this.#disposed) return;
		this.#disposed = true;
		this.#disposeInputs();
		this.#fontEvents?.removeEventListener('loadingdone', this.#fontsChanged);
		this.#fontEvents = undefined;
		this.#canvas.dispose();
		this.#unsubscribe();
		this.#eventUnsubscribe();
		this.controller.destroy();
		this.#root.replaceChildren();
	}
	connectedCallback(): void {
		if (this.#disposed) return;
		if (this.#suspended) this.#disposeInputs = this.#wireInputs();
		this.#suspended = false;
		this.#fontEvents = this.ownerDocument.fonts;
		this.#fontEvents?.addEventListener('loadingdone', this.#fontsChanged);
		this.#render(this.controller.state);
	}
	disconnectedCallback(): void {
		if (!this.#disposed) {
			this.#suspended = true;
			this.#disposeInputs();
			this.#fontEvents?.removeEventListener('loadingdone', this.#fontsChanged);
			this.#fontEvents = undefined;
			this.controller.cancelLoad();
			this.controller.cancelEdit();
			this.#canvas.dispose();
			this.#canvas.invalidate();
		}
	}
	#wireInputs(): () => void {
		const disposeChrome = this.#chrome.wire();
		const disposeCommands = this.#commands.wire();
		const disposeBackstage = this.#backstage.wire();
		const disposeMenus = wireContextMenus(this.#root, this.#viewport, this.controller);
		const disposeTellMe = wireTellMe(this.#root);
		const keyTips = attachKeyTips(this.#root);
		const disposePanZoom = this.#panZoom.wire();
		const disposeRulers = this.#rulers.wire();
		const disposeStencil = wireStencil(
			this.#root.querySelector('.shapes-pane')!,
			this.#viewport,
			this.controller,
			(message) => {
				this.#announcement = message;
				this.#status.textContent = message;
			},
		);
		const disposeFind = wireFindBar(this.#findBar, this.#search.input, () =>
			this.#viewport.focus({ preventScroll: true }),
		);
		const disposeEdit = this.#edit.wire();
		const disposeLayers = wireLayerControls(this.#layers, this.controller);
		const disposeInputs = wireViewerInputs(
			{
				viewport: this.#viewport,
				commandRoot: this.#root,
				zoomSlider: this.#zoomSlider,
				searchInput: this.#search.input,
			},
			this.controller,
			(mode) => this.#fit(mode),
		);
		return () => {
			disposeChrome();
			disposeCommands();
			disposeBackstage();
			disposeMenus();
			disposeTellMe();
			keyTips.dispose();
			disposePanZoom();
			disposeRulers();
			disposeStencil();
			disposeFind();
			disposeInputs();
			disposeLayers();
			disposeEdit();
		};
	}
	#assertAlive(): void {
		if (this.#disposed) throw new Error('The viewer has been destroyed.');
	}
	#render(state: ViewerState): void {
		if (this.#suspended) return;
		const page = state.document?.pages[state.pageIndex];
		const changed = this.#canvas.render(state);
		if (changed) this.#announcement = undefined;
		this.#zoomSlider.value = Math.round(state.zoom * 100);
		this.#zoomSlider.disabled = !page;
		const inspected = selectedShape(state.document, state.selectedShape, state.pageIndex);
		const inches = (value: number) => `${+value.toFixed(3)} in`;
		this.#shapeStatus.setAttribute(
			'value',
			inspected ? `Width: ${inches(inspected.width)}  Height: ${inches(inspected.height)}` : '',
		);
		for (const button of this.#root.querySelectorAll<HTMLButtonElement>('[data-action]'))
			button.disabled = !page;
		renderSearchControls(this.#search, state);
		this.#edit.render(state);
		if (changed) renderLayerControls(this.#layers, state);
		else
			this.#layers.querySelector<HTMLButtonElement>('[data-layer-reset="all"]')!.disabled =
				state.layerVisibilityOverrides.length === 0;
		this.#chrome.render(state, this.#notes.children.length);
		this.#commands.render(state);
		this.#backstage.render(state);
		this.#panZoom.render();
		this.#viewport.setAttribute('aria-busy', String(state.loading || state.edit.busy));
		this.#status.textContent = state.loading
			? 'Opening diagram…'
			: state.edit.busy
				? 'Updating diagram…'
				: (state.error?.message ??
					(state.edit.error ? `Edit rejected: ${editErrorMessage(state.edit.error)}` : undefined) ??
					this.#announcement ??
					(page ? `${page.name} · ${page.shapes.length} top-level shapes` : 'No diagram open'));
		const warnings = (state.document?.diagnostics.length ?? 0) + this.#canvas.warnings.length;
		this.#diagnostics.textContent = warnings
			? `${this.#notes.children.length} compatibility notes`
			: 'Local-only viewing';
	}
}
export function registerVisioViewer(): void {
	if (typeof customElements === 'undefined')
		throw new Error('Register the Visio viewer in a browser.');
	if (!customElements.get('visio-viewer'))
		customElements.define('visio-viewer', VisioViewerElement);
}
declare global {
	interface HTMLElementTagNameMap {
		'visio-viewer': VisioViewerElement;
	}
}
