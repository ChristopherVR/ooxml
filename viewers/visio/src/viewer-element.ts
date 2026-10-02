import type { VisioDocument, VisioShape } from 'ooxml-core/visio';
import type { VsdxSource } from './contract.js';
import { createWorkerParser } from './worker-parser.js';
import { ViewerController, type ViewerState } from './controller.js';
import { renderPage } from './render-svg.js';
import { MAX_INPUT_BYTES } from './scene-validation.js';
import { selectedShape, shapeDetails } from './shape-inspector.js';
import { compatibilityNotes, compatibilityText } from './diagnostics.js';
import { wireViewerInputs } from './viewer-input.js';
import {
	layerControlsTemplate,
	renderLayerControls,
	wireLayerControls,
} from './viewer-layer-controls.js';
import { viewerStyles } from './styles.js';
import {
	searchTemplate,
	searchControls,
	renderSearchControls,
	type SearchControls,
} from './viewer-search.js';
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
	#pageSelect: HTMLSelectElement;
	#zoomLabel: HTMLOutputElement;
	#status: HTMLSpanElement;
	#diagnostics: HTMLSpanElement;
	#toolbar: HTMLDivElement;
	#search: SearchControls;
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
		// This template is static, never document content.
		this.#root.innerHTML = `<style>${viewerStyles}</style><div class="toolbar" role="group" aria-label="Diagram controls"><label>Page <select aria-label="Page"></select></label><span class="spacer"></span><button type="button" data-action="out" aria-label="Zoom out">−</button><output class="zoom" aria-label="Zoom level">100%</output><button type="button" data-action="in" aria-label="Zoom in">+</button><button type="button" data-action="fit">Fit page</button><button type="button" data-action="actual">100%</button>${searchTemplate}${layerControlsTemplate}</div><div class="viewport" tabindex="0" role="region" aria-label="Diagram canvas"></div><details class="shape-inspector" hidden><summary>Selected shape</summary><div></div></details><details class="notes"><summary>Compatibility notes</summary><ul></ul></details><div class="status" role="status"><span data-status></span><span data-diagnostics></span></div>`;
		this.#viewport = this.#root.querySelector('.viewport')!;
		this.#pageSelect = this.#root.querySelector('select')!;
		this.#zoomLabel = this.#root.querySelector('output')!;
		this.#status = this.#root.querySelector('[data-status]')!;
		this.#diagnostics = this.#root.querySelector('[data-diagnostics]')!;
		this.#toolbar = this.#root.querySelector('.toolbar')!;
		this.#search = searchControls(this.#root);
		this.#layers = this.#root.querySelector('.layer-controls')!;
		this.#notes = this.#root.querySelector('.notes ul')!;
		this.#notesPanel = this.#root.querySelector('.notes')!;
		this.#inspector = this.#root.querySelector('.shape-inspector')!;
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
	}
	get renderWarnings(): readonly string[] {
		return this.#renderWarnings;
	}
	async load(source: VsdxSource): Promise<void> {
		this.#assertAlive();
		await this.controller.loadSource(() => {
			if ((source instanceof Blob ? source.size : source.byteLength) > MAX_INPUT_BYTES)
				throw new Error('This viewer accepts files up to 32 MiB.');
			return source instanceof Blob ? source.arrayBuffer() : source;
		});
	}
	fit(): void {
		this.#assertAlive();
		const page = this.document?.pages[this.pageIndex];
		if (!page) return;
		const width = Math.max(1, this.#viewport.clientWidth - 64);
		const height = Math.max(1, this.#viewport.clientHeight - 64);
		this.zoom = Math.min(width / (page.width * 96), height / (page.height * 96));
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
			this.#disposeRenderer();
			this.#document = null;
			this.#renderedPage = -1;
		}
	}
	#wireInputs(): () => void {
		const disposeLayers = wireLayerControls(this.#layers, this.controller);
		const disposeInputs = wireViewerInputs(
			{
				viewport: this.#viewport,
				toolbar: this.#toolbar,
				pageSelect: this.#pageSelect,
				searchInput: this.#search.input,
			},
			this.controller,
			() => this.fit(),
		);
		return () => {
			disposeInputs();
			disposeLayers();
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
			this.#pageSelect.replaceChildren();
			for (const [index, candidate] of (state.document?.pages ?? []).entries()) {
				const option = document.createElement('option');
				option.value = String(index);
				option.textContent = candidate.name + (candidate.isBackground ? ' (background)' : '');
				this.#pageSelect.append(option);
			}
			this.#pageSelect.value = String(state.pageIndex);
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
				heading.textContent = 'Your diagrams, on your terms';
				empty.append(heading, 'Open a .vsdx file to start. Files stay in this browser.');
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
				this.#inspector
					.querySelector('div')!
					.replaceChildren(...(inspected ? [shapeDetails(inspected)] : []));
			}
			this.#renderedSelection = selection ? { ...selection } : null;
		}
		if (!searchResult) this.#revealedSearchResult = undefined;
		this.#zoomLabel.value = `${Math.round(state.zoom * 100)}%`;
		this.#pageSelect.disabled = !page;
		for (const button of this.#toolbar.querySelectorAll('button')) button.disabled = !page;
		renderSearchControls(this.#search, state);
		if (changed) renderLayerControls(this.#layers, state);
		else
			this.#layers.querySelector<HTMLButtonElement>('[data-layer-reset="all"]')!.disabled =
				state.layerVisibilityOverrides.length === 0;
		this.#viewport.setAttribute('aria-busy', String(state.loading));
		this.#status.textContent = state.loading
			? 'Opening diagram…'
			: (state.error?.message ??
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
