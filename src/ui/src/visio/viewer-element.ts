import type { VisioDocument, VisioEdit } from 'ooxml-core/visio';
import type { VsdxSource, VisioShapeSelection } from 'ooxml-core/visio/ui';
import { createWorkerParser } from './worker-parser';
import { ViewerController, type ViewerState } from './controller';
import { MAX_INPUT_BYTES } from 'ooxml-core/visio/ui';
import { selectedShape } from './shape-inspector';
import { wireViewerInputs } from './viewer-input';
import { renderLayerControls, wireLayerControls } from './viewer-layer-controls';
import { viewerStyles } from './styles';
import { canvasAndRibbonStyles, visioThemeAliases } from './styles/index';
import { installOfficeUiTheme } from '../theme';
import { createRibbon } from './ribbon';
import { applyKeyTips } from './ribbon-keytips';
import { attachKeyTips } from '../controls';
import { createPageTabs, createStatusBar } from './status-bar';
import { createFindBar, renderFindBar, wireFindBar, type FindBar } from './viewer-search';
import { ViewerReplace } from './viewer-replace';
import { fitZoom } from './viewer-fit';
import { createShapesStrip, createShapesWindow } from './shapes-window';
import { createBackstage, type BackstagePage } from './backstage';
import type { CreateVsdxOptions } from 'ooxml-core/visio';
import { ViewerBackstage } from './viewer-backstage';
import { ViewerPointerGestures } from './viewer-pointer-gestures';
import { createSizePosition, ViewerSizePosition } from './viewer-size-position';
import { createContextMenus, wireContextMenus } from './viewer-context-menu';
import { wireTellMe } from './viewer-tell-me';
import { createPanZoom, ViewerPanZoom } from './viewer-pan-zoom';
import { createOptionsDialog, ViewerProfile } from './viewer-options';
import { ViewerShare } from './viewer-share';
import { wireStencil } from './viewer-stencil';
import { createRulers, type Rulers } from './viewer-ruler';
import { ViewerEditControls } from './viewer-edit-controls';
import { ViewerChrome, viewerChromeTemplate } from './viewer-chrome';
import { ViewerCanvas } from './viewer-canvas';
import { ViewerCommands } from './viewer-commands';
import { ViewerPresentation } from './viewer-presentation';
import { ViewerLineEndpoints } from './viewer-line-endpoints';
import { ViewerRotationHandle } from './viewer-rotation-handle';
import { ViewerResizeHandles } from './viewer-resize-handles';
import { editErrorMessage } from 'ooxml-core/visio/ui';
import { registerViewerControls } from './office-ui';
import { exportPageSvg, type SvgExportOptions, type SvgExportResult } from './export-svg';
import { exportDocumentPdf, exportPagePng } from './viewer-export';
import {
	createPrintSnapshot,
	type CurrentPagePrintSnapshotOptions,
	type PrintSnapshot,
} from './print-snapshot';

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
	#presentation: ViewerPresentation;
	#lineEndpoints: ViewerLineEndpoints;
	#rotationHandle: ViewerRotationHandle;
	#resizeHandles: ViewerResizeHandles;
	#handleState: ViewerState | undefined;
	#rulers: Rulers;
	#canvas: ViewerCanvas;
	#panZoom: ViewerPanZoom;
	#profile: ViewerProfile;
	#share: ViewerShare;
	#backstage: ViewerBackstage;
	#pointer: ViewerPointerGestures;
	#sizePosition: ViewerSizePosition;
	#fileName = '';
	#loadToken = 0;
	#findBar: FindBar;
	#replace: ViewerReplace;
	#status: HTMLSpanElement;
	#diagnostics: HTMLSpanElement;
	#toolbar: HTMLDivElement;
	#chrome: ViewerChrome;
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
		installOfficeUiTheme();
		// This template is static, never document content.
		this.#root.innerHTML = `<style>${visioThemeAliases}${viewerStyles}${canvasAndRibbonStyles}</style>${viewerChromeTemplate}`;
		const workspace = this.#root.querySelector('.workspace')!;
		this.#findBar = createFindBar(document);
		const ribbon = createRibbon(document);
		applyKeyTips(ribbon);
		workspace.before(ribbon, this.#findBar);
		workspace.prepend(createShapesStrip(document), createShapesWindow(document));
		workspace.append(createPanZoom(document), createSizePosition(document));
		this.#root.append(
			createBackstage(document),
			createOptionsDialog(document),
			...createContextMenus(document),
		);
		workspace.after(createPageTabs(document), createStatusBar(document));
		this.#viewport = this.#root.querySelector('.viewport')!;
		this.#rulers = createRulers(this.#viewport);
		this.#zoomSlider = this.#root.querySelector('office-ui-zoom-slider')!;
		this.#shapeStatus = this.#root.querySelector('[data-shape-status]')!;
		this.#status = this.#root.querySelector('[data-status]')!;
		this.#diagnostics = this.#root.querySelector('[data-diagnostics]')!;
		this.#toolbar = this.#root.querySelector('.toolbar')!;
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
		this.#sizePosition = new ViewerSizePosition(this.#root, this.controller, (open) =>
			this.#root.querySelector('[command="size-position"]')?.setAttribute('checked', String(open)),
		);
		this.#replace = new ViewerReplace(this.#findBar, this.controller, (message) => {
			this.#announcement = message;
			this.#status.textContent = message;
		});
		this.#presentation = new ViewerPresentation(this.#root, this.controller);
		this.#commands = new ViewerCommands({
			root: this.#root,
			viewport: this.#viewport,
			controller: this.controller,
			toolChanged: () => {
				this.#pointer.render(this.controller.state);
				this.#lineEndpoints.render(this.controller.state);
				this.#rotationHandle.render(this.controller.state);
				this.#resizeHandles.render(this.controller.state);
			},
			fit: (mode) => this.#fit(mode),
			togglePane: (pane) => this.#chrome.togglePane(pane),
			reveal: (panel, focusText) => this.#chrome.reveal(panel, focusText),
			rulers: this.#rulers,
			togglePanZoom: () => this.#panZoom.toggle(),
			present: () => this.#presentation.start(),
			presenting: () => this.#presentation.active,
			toggleSizePosition: () => this.#sizePosition.toggle(),
			focusSearch: () => {
				this.#chrome.closeCompactTools();
				this.#replace.showFind();
			},
			focusReplace: () => {
				this.#chrome.closeCompactTools();
				this.#replace.showReplace();
			},
			announce: (message) => {
				this.#announcement = message;
				this.#status.textContent = message;
			},
		});
		this.#profile = new ViewerProfile(this.#root);
		this.#pointer = new ViewerPointerGestures(this.#viewport, this.controller, {
			active: () => this.#commands.tool === 'pointer',
			announce: (message) => {
				this.#announcement = message;
				this.#status.textContent = message;
			},
			snap: (page, ids, delta) =>
				this.#commands.layoutCommands.guides.snap(page, ids, delta, this.#commands.gridStep),
			clearSnap: () => this.#commands.layoutCommands.guides.clearHints(),
		});
		this.#lineEndpoints = new ViewerLineEndpoints(this.#viewport, this.controller, {
			active: () => this.#commands.tool === 'pointer',
			announce: (message) => {
				this.#announcement = message;
				this.#status.textContent = message;
			},
		});
		this.#rotationHandle = new ViewerRotationHandle(this.#viewport, this.controller, {
			active: () => this.#commands.tool === 'pointer',
			announce: (message) => {
				this.#announcement = message;
				this.#status.textContent = message;
			},
		});
		this.#resizeHandles = new ViewerResizeHandles(this.#viewport, this.controller, {
			active: () => this.#commands.tool === 'pointer',
			announce: (message) => {
				this.#announcement = message;
				this.#status.textContent = message;
			},
		});
		this.#share = new ViewerShare(this.#root, this.controller, () => this.#profile.profile);
		this.#backstage = new ViewerBackstage({
			root: this.#root,
			viewport: this.#viewport,
			fileName: () => this.#fileName,
			load: (source) => this.load(source),
			createBlankDrawing: () => this.createBlankDrawing(),
			exportVsdx: () => this.controller.exportVsdx(),
			exportSvg: () => this.exportSvg(),
			exportPicture: async (format) => {
				const { document: model, pageIndex } = this.controller.state;
				if (!model) throw new Error('Open a drawing before exporting.');
				const doc = this.ownerDocument;
				return {
					blob:
						format === 'pdf'
							? await exportDocumentPdf(doc, model)
							: await exportPagePng(doc, model, pageIndex),
					pageIndex,
				};
			},
			closeDocument: () => {
				this.#fileName = '';
				this.controller.setDocument(null);
			},
			revealNotes: () => this.#chrome.reveal('notes'),
			showOptions: () => this.#profile.showOptions(),
			announce: (message) => {
				this.#announcement = message;
				this.#status.textContent = message;
			},
			noteCount: () => this.#notes.children.length,
			pageSetup: () =>
				this.#commands.run({ type: 'page-setup', command: { op: 'dialog', tab: 'print' } }),
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
	/** Create a source-backed drawing and name only the accepted request. */
	async createBlankDrawing(options?: CreateVsdxOptions): Promise<void> {
		this.#assertAlive();
		const token = ++this.#loadToken;
		const generation = this.controller.documentGeneration;
		await this.controller.createBlankDrawing(options);
		if (
			this.#disposed ||
			token !== this.#loadToken ||
			this.controller.documentGeneration !== generation + 1 ||
			this.controller.state.loading ||
			!this.controller.state.edit.sourceAvailable
		)
			return;
		this.#fileName = 'New drawing.vsdx';
		this.#backstage.hide();
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
	selectShapes(shapes: readonly VisioShapeSelection[]): void {
		this.#assertAlive();
		this.controller.selectShapes(shapes);
	}
	selectAll(): void {
		this.#assertAlive();
		this.controller.selectAll();
	}
	clearSelection(): void {
		this.#assertAlive();
		this.controller.clearSelection();
	}
	duplicateSelection(): Promise<void> {
		this.#assertAlive();
		return this.controller.duplicateSelection();
	}
	copySelection(): Promise<void> {
		this.#assertAlive();
		return this.#commands.clipboard('copy');
	}
	cutSelection(): Promise<void> {
		this.#assertAlive();
		return this.#commands.clipboard('cut');
	}
	pasteSelection(): Promise<void> {
		this.#assertAlive();
		return this.#commands.clipboard('paste');
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
	/** Start Visio's Presentation Mode on the current foreground page (F5). */
	startPresentation(): void {
		this.#assertAlive();
		this.#presentation.start();
	}
	/** Leave Presentation Mode; the drawing window's page, zoom and selection are unchanged. */
	exitPresentation(): void {
		this.#assertAlive();
		this.#presentation.exit();
	}
	get presenting(): boolean {
		return this.#presentation.active;
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
		const disposePresentation = this.#presentation.wire();
		const disposePointer = this.#pointer.wire();
		const disposeLineEndpoints = this.#lineEndpoints.wire();
		const disposeRotation = this.#rotationHandle.wire();
		const disposeResize = this.#resizeHandles.wire();
		const disposeBackstage = this.#backstage.wire();
		const disposeMenus = wireContextMenus(this.#root, this.#viewport, this.controller);
		const disposeTellMe = wireTellMe(this.#root);
		const keyTips = attachKeyTips(this.#root);
		const disposePanZoom = this.#panZoom.wire();
		const disposeSizePosition = this.#sizePosition.wire();
		const disposeProfile = this.#profile.wire();
		const disposeShare = this.#share.wire();
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
		const disposeFind = wireFindBar(this.#findBar, this.controller, () =>
			this.#viewport.focus({ preventScroll: true }),
		);
		const disposeReplace = this.#replace.wire();
		const disposeEdit = this.#edit.wire();
		const disposeLayers = wireLayerControls(this.#layers, this.controller);
		const disposeInputs = wireViewerInputs(
			{
				viewport: this.#viewport,
				zoomSlider: this.#zoomSlider,
			},
			this.controller,
			(mode) => this.#fit(mode),
			(message) => {
				this.#announcement = message;
				this.#status.textContent = message;
			},
		);
		return () => {
			disposeChrome();
			disposeCommands();
			disposePresentation();
			disposePointer();
			disposeLineEndpoints();
			disposeRotation();
			disposeResize();
			disposeBackstage();
			disposeMenus();
			disposeTellMe();
			keyTips.dispose();
			disposePanZoom();
			disposeSizePosition();
			disposeProfile();
			disposeShare();
			disposeRulers();
			disposeStencil();
			disposeFind();
			disposeReplace();
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
		const previous = this.#handleState;
		this.#handleState = state;
		if (
			changed ||
			!previous ||
			previous.zoom !== state.zoom ||
			previous.selectedShape !== state.selectedShape ||
			previous.selectedShapes !== state.selectedShapes ||
			previous.loading !== state.loading ||
			previous.edit.busy !== state.edit.busy ||
			previous.edit.sourceAvailable !== state.edit.sourceAvailable
		) {
			this.#lineEndpoints.render(state);
			this.#rotationHandle.render(state);
			this.#resizeHandles.render(state);
		}
		if (changed) this.#announcement = undefined;
		this.#zoomSlider.value = Math.round(state.zoom * 100);
		this.#zoomSlider.disabled = !page;
		const inspected = selectedShape(state.document, state.selectedShape, state.pageIndex);
		const inches = (value: number) => `${+value.toFixed(3)} in`;
		this.#shapeStatus.setAttribute(
			'value',
			state.selectedShapes.length > 1
				? `${state.selectedShapes.length} shapes selected`
				: inspected
					? `Width: ${inches(inspected.width)}  Height: ${inches(inspected.height)}`
					: '',
		);
		for (const button of this.#root.querySelectorAll<HTMLButtonElement>('[data-action]'))
			button.disabled = !page;
		renderFindBar(this.#findBar, state);
		this.#replace.render(state);
		this.#edit.render(state);
		if (changed) renderLayerControls(this.#layers, state);
		else
			this.#layers.querySelector<HTMLButtonElement>('[data-layer-reset="all"]')!.disabled =
				state.layerVisibilityOverrides.length === 0;
		this.#chrome.render(state, this.#notes.children.length);
		this.#commands.render(state);
		this.#presentation.render(state);
		this.#backstage.render(state);
		this.#pointer.render(state);
		this.#panZoom.render();
		this.#sizePosition.render(state);
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
