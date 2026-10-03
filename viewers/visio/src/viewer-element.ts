import type { VisioDocument, VisioShape, VisioEdit } from 'ooxml-core/visio';
import type { VsdxSource } from './contract.js';
import { createWorkerParser } from './worker-parser.js';
import { ViewerController, type ViewerState } from './controller.js';
import { renderPage } from './render-svg.js';
import { MAX_INPUT_BYTES } from './scene-validation.js';
import { selectedShape, shapeDetails } from './shape-inspector.js';
import { compatibilityNotes, compatibilityText } from './diagnostics.js';
import { wireViewerInputs } from './viewer-input.js';
import { renderLayerControls, wireLayerControls } from './viewer-layer-controls.js';
import { viewerStyles } from './styles.js';
import { canvasAndRibbonStyles } from './styles/index.js';
import { createRibbon } from './ribbon.js';
import { createPageTabs, createStatusBar } from './status-bar.js';
import { createFindBar, wireFindBar } from './viewer-search.js';
import { fitZoom } from './viewer-fit.js';
import { createShapesWindow } from './shapes-window.js';
import { wireStencil } from './viewer-stencil.js';
import { createRulers, type Rulers } from './viewer-ruler.js';
import { ViewerEditControls } from './viewer-edit-controls.js';
import { searchControls, renderSearchControls, type SearchControls } from './viewer-search.js';
import { ViewerChrome, viewerChromeTemplate } from './viewer-chrome.js';
import { ViewerCommands } from './viewer-commands.js';
import { editErrorMessage } from './edit-error.js';
import { registerViewerControls } from './office-ui.js';
import type { TextSearchResult } from './document-text-search.js';
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
	#findBar: HTMLElement;
	#status: HTMLSpanElement;
	#diagnostics: HTMLSpanElement;
	#toolbar: HTMLDivElement;
	#chrome: ViewerChrome;
	#search: SearchControls;
	#edit: ViewerEditControls;
	#layers: HTMLDetailsElement;
	#renderedLayerOverrides: ViewerState['layerVisibilityOverrides'] | undefined;
	#revealedSearchResult: TextSearchResult | undefined;
	#notes: HTMLUListElement;
	#notesPanel: HTMLDetailsElement;
	#inspector: HTMLDetailsElement;
	#inspectedShape: VisioShape | undefined;
	#renderedSelection: ViewerState['selectedShape'] = null;
	#unsubscribe: () => void;
	#eventUnsubscribe: () => void;
	#document: VisioDocument | null = null;
	#renderedPage = -1;
	#disposed = false;
	#suspended = false;
	#renderWarnings: string[] = [];
	#disposeRenderer: () => void = () => {};
	#fontEvents: FontFaceSet | undefined;
	#fontsChanged = () => {
		if (this.#disposed) return;
		this.#document = null;
		this.#renderedPage = -1;
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
		workspace.before(createRibbon(document), this.#findBar);
		workspace.prepend(createShapesWindow(document));
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
		this.#notesPanel = this.#root.querySelector('.notes')!;
		this.#inspector = this.#root.querySelector('.shape-inspector')!;
		this.#chrome = new ViewerChrome(this.#root, this.controller);
		this.#commands = new ViewerCommands({
			root: this.#root,
			viewport: this.#viewport,
			controller: this.controller,
			fit: (mode) => this.#fit(mode),
			togglePane: (pane) => this.#chrome.togglePane(pane),
			reveal: (panel, focusText) => this.#chrome.reveal(panel, focusText),
			rulers: this.#rulers,
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
		return this.#renderWarnings;
	}
	async load(source: VsdxSource): Promise<void> {
		this.#assertAlive();
		if (source instanceof Blob) {
			await this.controller.loadSource(() => {
				if (source.size > MAX_INPUT_BYTES)
					throw new Error('This viewer accepts files up to 32 MiB.');
				return source.arrayBuffer();
			});
		} else await this.controller.load(source);
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
		this.#disposeRenderer();
		this.#unsubscribe();
		this.#eventUnsubscribe();
		this.controller.destroy();
		this.#revealedSearchResult = undefined;
		this.#renderedSelection = null;
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
			this.#disposeRenderer();
			this.#document = null;
			this.#renderedPage = -1;
		}
	}
	#wireInputs(): () => void {
		const disposeChrome = this.#chrome.wire();
		const disposeCommands = this.#commands.wire();
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
		const changed =
			this.#document !== state.document ||
			this.#renderedPage !== state.pageIndex ||
			this.#renderedLayerOverrides !== state.layerVisibilityOverrides;
		if (changed) {
			this.#disposeRenderer();
			this.#disposeRenderer = () => {};
			this.#document = state.document;
			this.#renderedPage = state.pageIndex;
			this.#renderedLayerOverrides = state.layerVisibilityOverrides;
			this.#announcement = undefined;
			this.#renderWarnings = [];
			if (state.document && page) {
				const result = renderPage(state.document, page, {
					layerVisibilityOverrides: state.layerVisibilityOverrides,
				});
				this.#renderWarnings = result.warnings;
				this.#disposeRenderer = result.dispose;
				this.#viewport.replaceChildren(result.svg);
			} else {
				const empty = document.createElement('div');
				empty.className = 'empty';
				const heading = document.createElement('strong');
				heading.textContent = 'Open a Visio drawing';
				empty.append(
					heading,
					'Open a .vsdx or supported .vsd file to start. Files stay in this browser.',
				);
				this.#viewport.replaceChildren(empty);
			}
			const notes = compatibilityNotes(state.document?.diagnostics ?? [], this.#renderWarnings);
			this.#notes.replaceChildren(
				...notes.map((note) => {
					const item = document.createElement('li');
					item.textContent = compatibilityText(note);
					return item;
				}),
			);
			this.#notesPanel.hidden = !notes.length;
		}
		const svg = this.#viewport.querySelector<SVGSVGElement>('svg');
		if (svg && page) {
			svg.style.width = `${page.width * 96 * state.zoom}px`;
			svg.style.height = `${page.height * 96 * state.zoom}px`;
			// Viewer-only grid spacing: quarter-inch minor and one-inch major lines.
			svg.style.setProperty('--vv-inch', `${96 * state.zoom}px`);
		}
		const searchResult = state.search.results[state.search.activeIndex];
		const selection = state.selectedShape;
		if (
			changed ||
			this.#renderedSelection?.id !== selection?.id ||
			this.#renderedSelection?.pageId !== selection?.pageId ||
			this.#renderedSelection?.name !== selection?.name ||
			(searchResult && searchResult !== this.#revealedSearchResult)
		) {
			for (const shape of this.#viewport.querySelectorAll<SVGGElement>('[data-shape-id]')) {
				const selected =
					!!state.selectedShape &&
					shape.dataset.shapeId === state.selectedShape.id &&
					(!state.selectedShape.pageId || shape.dataset.pageId === state.selectedShape.pageId);
				shape.dataset.selected = String(selected);
				if (
					selected &&
					searchResult &&
					searchResult !== this.#revealedSearchResult &&
					shape.dataset.shapeId === searchResult.shapeId &&
					shape.dataset.pageId === searchResult.pageId
				) {
					shape.scrollIntoView?.({ block: 'nearest', inline: 'nearest' });
					this.#revealedSearchResult = searchResult;
				}
				if (state.selectedShape) shape.setAttribute('tabindex', selected ? '0' : '-1');
				if (shape.getAttribute('role') === 'button')
					shape.setAttribute('aria-pressed', String(selected));
			}
			const inspected = selectedShape(state.document, state.selectedShape, state.pageIndex);
			if (inspected !== this.#inspectedShape) {
				this.#inspectedShape = inspected;
				this.#inspector.hidden = !inspected;
				if (inspected) this.#inspector.open = true;
				this.#inspector
					.querySelector('div')!
					.replaceChildren(...(inspected ? [shapeDetails(inspected)] : []));
			}
			this.#renderedSelection = selection ? { ...selection } : null;
		}
		if (!searchResult) this.#revealedSearchResult = undefined;
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
		this.#viewport.setAttribute('aria-busy', String(state.loading || state.edit.busy));
		this.#status.textContent = state.loading
			? 'Opening diagram…'
			: state.edit.busy
				? 'Updating diagram…'
				: (state.error?.message ??
					(state.edit.error ? `Edit rejected: ${editErrorMessage(state.edit.error)}` : undefined) ??
					this.#announcement ??
					(page ? `${page.name} · ${page.shapes.length} top-level shapes` : 'No diagram open'));
		const warnings = (state.document?.diagnostics.length ?? 0) + this.#renderWarnings.length;
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
