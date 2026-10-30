import type { VisioDocument, VisioEdit } from 'ooxml-core/visio';
import type { VsdxSource, VisioShapeSelection } from 'ooxml-core/visio/ui';
import { createWorkerParser } from './worker-parser';
import { ViewerController, type ViewerState } from './controller';
import { MAX_INPUT_BYTES } from 'ooxml-core/visio/ui';
import { selectedShape } from './shape-inspector';
import { wireViewerInputs } from './viewer-input';
import { viewerStyles } from './styles';
import { canvasAndRibbonStyles, visioThemeAliases } from './styles/index';
import { installOfficeUiThemeFor } from '../theme';
import { createRibbon } from './ribbon';
import { wireInlineGalleries } from './ribbon-inline-galleries';
import { syncRibbonAddIns, type RibbonAddInTab } from '../ribbon/add-in-tabs';
import { commandRow } from './ribbon-parts';
import { applyKeyTips } from './ribbon-keytips';
import { attachKeyTips } from '../controls';
import { createPageTabs, createStatusBar, statusLanguage } from './status-bar';
import { createTitleBar, renderTitleBar } from './title-bar';
import { localizeRibbon } from './ribbon-localize';
import { applyScreenTips, restoreScreenTips } from './viewer-screen-tips';
import { CUSTOMIZABLE_TABS, DEFAULT_BEHAVIOUR, type VisioBehaviour } from './viewer-behaviour';
import { wireRibbonOverflow } from './ribbon-overflow';
import { renderContextualTabs } from './viewer-contextual-tabs';
import { wireRibbonTabCue } from './ribbon-tab-cue';
import { createFindBar, renderFindBar, wireFindBar, type FindBar } from './viewer-search';
import { ViewerReplace } from './viewer-replace';
import { fitZoom } from './viewer-fit';
import { createShapesStrip, createShapesWindow } from './shapes-window';
import { createBackstage, type BackstagePage } from './backstage';
import type { CreateVsdxOptions } from 'ooxml-core/visio';
import type { VisioTemplateId } from 'ooxml-core/visio/ui';
import { ViewerBackstage } from './viewer-backstage';
import type { ViewerPointerGestures } from './viewer-pointer-gestures';
import { createSizePosition, ViewerSizePosition } from './viewer-size-position';
import { ViewerInlineSize } from './viewer-inline-size';
import { createContextMenus, wireContextMenus } from './viewer-context-menu';
import { syncShapeActions } from './viewer-shape-actions';
import {
	applyVisioCommandVisibility,
	quickAccessShown,
	syncVisioHostMenu,
	ViewerCustomizationState,
	visioActionBlocked,
} from './viewer-customization';
import type { VisioCustomization } from 'ooxml-core/visio/ui';
import { ViewerMiniToolbar } from './viewer-mini-toolbar';
import { ViewerLivePreview } from './viewer-live-preview';
import {
	ViewerAutoRecover,
	indexedDbRecoveryStore,
	storageRecoveryStore,
} from './viewer-autorecover';
import type { DocumentStorage } from 'ooxml-core/embedding';
import { announceFileCommand } from './viewer-file-commands';
import { RecoverDialog } from './viewer-recover-dialog';
import { wireTellMe } from './viewer-tell-me';
import { createPanZoom, ViewerPanZoom } from './viewer-pan-zoom';
import {
	createOptionsDialog,
	readStatusBarTheme,
	ViewerProfile,
	type StatusBarTheme,
} from './viewer-options';
import { ViewerShare } from './viewer-share';
import type { VisioCollaborationOptions, VisioCollaborationState } from './collaboration-types';
import { wireStencil } from './viewer-stencil';
import { createRulers, type Rulers } from './viewer-ruler';
import type { ViewerInlineText } from './viewer-inline-text';
import { createCanvasEditing } from './canvas-editing';
import { ViewerChrome, viewerChromeTemplate } from './viewer-chrome';
import { ViewerCanvas } from './viewer-canvas';
import { ViewerCommands } from './viewer-commands';
import { ViewerPresentation } from './viewer-presentation';
import type { ViewerLineEndpoints } from './viewer-line-endpoints';
import type { ViewerControlHandles } from './viewer-control-handles';
import type { ViewerRotationHandle } from './viewer-rotation-handle';
import type { ViewerResizeHandles } from './viewer-resize-handles';
import { editErrorMessage } from 'ooxml-core/visio/ui';
import { registerViewerControls } from './office-ui';
import { exportPageSvg, type SvgExportOptions, type SvgExportResult } from './export-svg';
import { exportPagePng } from './viewer-export';
import { exportDocumentPdfDetailed } from './viewer-pdf';
import { pdfOptionsFor } from './viewer-pdf-options';
import type { PdfFontFile } from '../export/pdf-embedded-font';
import {
	createPrintSnapshot,
	type CurrentPagePrintSnapshotOptions,
	type PrintSnapshot,
} from './print-snapshot';

const BaseElement = (
	typeof HTMLElement === 'undefined' ? class {} : HTMLElement
) as typeof HTMLElement;
export class VisioViewerElement extends BaseElement {
	/** False while the drawing cannot be edited: `canEdit`, `readOnly`, a `viewer` role or `capabilities.edit`. */
	get canEdit(): boolean {
		return this.controller.canEdit;
	}
	set canEdit(value: boolean) {
		this.controller.canEdit = value;
	}
	/**
	 * The host customisation (`VisioCustomization`): `capabilities`, `backstage.extraPages` and
	 * `contextMenu.extraCommands`. Applied when the object identity changes; null restores the
	 * defaults.
	 */
	get customization(): VisioCustomization | null {
		return this.#custom.input;
	}
	set customization(value: VisioCustomization | null | undefined) {
		this.#assertAlive();
		if (!this.#custom.set(value ?? null)) return;
		this.controller.setCapabilityReadOnly(!this.#custom.resolved.canEdit);
		applyVisioCommandVisibility(this.#root, this.#custom.resolved);
		this.#backstage.refreshCustomization();
		renderTitleBar(
			this.#titleBar,
			this.controller.state,
			this.#fileName,
			this.#profile.profile,
			quickAccessShown(this.#behaviour, this.#custom.resolved),
		);
	}
	readonly #custom = new ViewerCustomizationState();
	/**
	 * The inverse of `canEdit`: true makes the drawing read-only and wins over `canEdit`. A
	 * collaboration `viewer` role also reads as read-only, and turning this off never lifts it.
	 */
	get readOnly(): boolean {
		return this.controller.readOnly;
	}
	set readOnly(value: boolean) {
		this.controller.readOnly = value;
	}
	/**
	 * Shares the open drawing in a room (the shared embedding contract of `ooxml-core/collab`):
	 * while set, the drawing is shared and one loaded later rejoins; null leaves the room.
	 */
	get collaboration(): VisioCollaborationOptions | null {
		return this.#share.wanted;
	}
	set collaboration(value: VisioCollaborationOptions | null | undefined) {
		this.#share.want(value ?? null);
	}
	/**
	 * Host-owned storage (`ooxml-core/embedding`). When set, File > Save and the Quick Access Save
	 * hand the drawing to `storage.save()` and mark it saved instead of downloading a copy, File >
	 * Open asks `storage.open()` when it has one, and AutoRecover keeps its copies through the
	 * storage's recovery methods (when it has all three) instead of IndexedDB. The cancelable
	 * `file-command` event still fires first.
	 */
	get storage(): DocumentStorage | null {
		return this.#storage ?? null;
	}
	set storage(value: DocumentStorage | null | undefined) {
		this.#storage = value ?? undefined;
	}
	/**
	 * TrueType fonts File > Export > PDF embeds (as subsets) so text outside the standard fonts'
	 * character set, or in a particular typeface, stays real text. The host supplies the bytes
	 * (nothing is bundled); a font whose licence forbids embedding is reported and skipped.
	 */
	pdfFonts: readonly PdfFontFile[] = [];
	/** What `collaboration-change` last reported. */
	get collaborationState(): VisioCollaborationState {
		return this.#share.state;
	}
	/** Leaves the room and clears `collaboration`. A host-owned session stays alive. */
	stopCollaboration(): void {
		this.#share.stop();
	}
	/** Restarts the connection, keeping the shared drawing and queued offline changes. */
	reconnectCollaboration(): void {
		this.#share.reconnect();
	}
	static observedAttributes = ['lang', 'locale'];
	#overflow: { refit(): void } | undefined;
	readonly controller = new ViewerController(createWorkerParser());
	#root: ShadowRoot;
	#disposeInputs: () => void;
	#viewport: HTMLDivElement;
	#zoomSlider: HTMLElement & { value: number; disabled: boolean };
	#announcement: string | undefined;
	#commands: ViewerCommands;
	#presentation: ViewerPresentation;
	#lineEndpoints: ViewerLineEndpoints;
	#controlHandles: ViewerControlHandles;
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
	#inlineSize: ViewerInlineSize;
	#fileName = '';
	#storage: DocumentStorage | undefined;
	#loadToken = 0;
	#findBar: FindBar;
	#replace: ViewerReplace;
	#status: HTMLSpanElement;
	#titleBar: HTMLElement;
	#toolbar: HTMLDivElement;
	#ribbonAddIns: readonly RibbonAddInTab[] = [];
	#chrome: ViewerChrome;
	#edit: ViewerInlineText;
	#notes: HTMLUListElement;
	#unsubscribe: () => void;
	#eventUnsubscribe: () => void;
	#disposed = false;
	#behaviour: VisioBehaviour = DEFAULT_BEHAVIOUR;
	#mini: ViewerMiniToolbar;
	#livePreview: ViewerLivePreview;
	#recover: ViewerAutoRecover;
	/** A drawing shown in place of the real one while a gallery tile is previewed. */
	#preview: { document: VisioDocument; base: VisioDocument | null } | undefined;
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
		installOfficeUiThemeFor(this);
		// This template is static, never document content.
		this.#root.innerHTML = `<style>${visioThemeAliases}${viewerStyles}${canvasAndRibbonStyles}</style>${viewerChromeTemplate}`;
		const workspace = this.#root.querySelector('.workspace')!;
		this.#findBar = createFindBar(document);
		const ribbon = createRibbon(document);
		applyKeyTips(ribbon);
		this.#titleBar = createTitleBar(document);
		workspace.before(this.#titleBar, ribbon, this.#findBar);
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
		this.#status = this.#root.querySelector('[data-status]')!;
		this.#toolbar = this.#root.querySelector('.toolbar')!;
		const say = (message: string): void => {
			this.#announcement = message;
			this.#status.textContent = message;
		};
		const recoverDialog = new RecoverDialog(this.#root, say, () => this.#locale());
		const browserStore = indexedDbRecoveryStore();
		this.#recover = new ViewerAutoRecover(
			this.controller,
			() => this.#fileName,
			() => storageRecoveryStore(this.#storage) ?? browserStore,
			(offer) => recoverDialog.show(offer),
		);
		this.#mini = new ViewerMiniToolbar(this.#root, this.#toolbar);
		this.#livePreview = new ViewerLivePreview(
			this.#root,
			this.controller,
			(document) => {
				this.#preview = { document, base: this.controller.state.document };
				this.#render(this.controller.state);
			},
			() => {
				if (!this.#preview) return;
				this.#preview = undefined;
				this.#render(this.controller.state);
			},
		);
		const announce = (message: string): void => {
			this.#announcement = message;
			this.#status.textContent = message;
		};
		// The same gestures the standalone building-block canvas uses; here the ribbon owns the tool.
		const editing = createCanvasEditing(this.#viewport, this.controller, {
			pointerActive: () => this.#commands.tool === 'pointer',
			announce,
			snap: (page, ids, delta, axis) =>
				this.#commands.layoutCommands.guides.snap(page, ids, delta, this.#commands.gridStep, axis),
			clearSnap: () => this.#commands.layoutCommands.guides.clearHints(),
			resizeSnap: {
				settings: () => this.#commands.layoutCommands.guides.resizeSnapSettings(),
				hint: (page, lines, matches) =>
					this.#commands.layoutCommands.guides.resizeHints(page, lines, matches),
				clear: () => this.#commands.layoutCommands.guides.clearHints(),
			},
		});
		this.#edit = editing.edit;
		this.#pointer = editing.pointer;
		this.#lineEndpoints = editing.lineEndpoints;
		this.#controlHandles = editing.controls;
		this.#rotationHandle = editing.rotation;
		this.#resizeHandles = editing.resize;
		// Selecting text in the editor shows the Mini Toolbar over it; collapsing the selection hides it.
		this.#edit.onSelection = (at) => (at ? this.#mini.showAt(at.x, at.y) : this.#mini.hide());
		this.#notes = this.#root.querySelector('.backstage-notes ul')!;
		this.#canvas = new ViewerCanvas(
			this.#viewport,
			this.#notes,
			this.#root.querySelector('.backstage-notes')!,
			this.#root.querySelector('.shape-inspector')!,
		);
		this.#chrome = new ViewerChrome(this.#root, this.controller);
		this.#panZoom = new ViewerPanZoom(this.#root, this.#viewport, this.controller, (open) =>
			this.#root.querySelector('[command="pan-zoom"]')?.setAttribute('checked', String(open)),
		);
		this.#inlineSize = new ViewerInlineSize(this.#root, this.controller);
		this.#sizePosition = new ViewerSizePosition(this.#root, this.controller, (open) =>
			this.#root.querySelector('[command="size-position"]')?.setAttribute('checked', String(open)),
		);
		this.#replace = new ViewerReplace(this.#findBar, this.controller, (message) => {
			this.#announcement = message;
			this.#status.textContent = message;
		});
		this.#presentation = new ViewerPresentation(this.#root, this.controller);
		this.#commands = new ViewerCommands({
			blocked: (action) => visioActionBlocked(this.#custom.resolved, action),
			root: this.#root,
			viewport: this.#viewport,
			controller: this.controller,
			toolChanged: () => {
				this.#pointer.render(this.controller.state);
				this.#lineEndpoints.render(this.controller.state);
				this.#controlHandles.render(this.controller.state);
				this.#rotationHandle.render(this.controller.state);
				this.#resizeHandles.render(this.controller.state);
			},
			fit: (mode) => this.#fit(mode),
			togglePane: (pane) => this.#chrome.togglePane(pane),
			reveal: (panel) => this.#reveal(panel),
			rulers: this.#rulers,
			togglePanZoom: () => this.#panZoom.toggle(),
			present: () => this.#presentation.start(),
			presenting: () => this.#presentation.active,
			toggleSizePosition: () => this.#sizePosition.toggle(),
			focusSearch: () => {
				this.#chrome.closeRibbonPopups();
				this.#replace.showFind();
			},
			focusReplace: () => {
				this.#chrome.closeRibbonPopups();
				this.#replace.showReplace();
			},
			announce: (message) => {
				this.#announcement = message;
				this.#status.textContent = message;
			},
		});
		const profileChanged = (): void =>
			renderTitleBar(
				this.#titleBar,
				this.controller.state,
				this.#fileName,
				this.#profile.profile,
				quickAccessShown(this.#behaviour, this.#custom.resolved),
			);
		this.#profile = new ViewerProfile(
			this.#root,
			profileChanged,
			{
				get: () => (this.#dark() ? 'dark' : this.statusBar),
				set: (theme) => {
					if (theme !== 'dark') this.statusBar = theme;
					if ((theme === 'dark') !== this.#dark())
						this.#requestScheme(theme === 'dark' ? 'dark' : 'light');
				},
			},
			(behaviour) => {
				this.#behaviour = behaviour;
				this.#applyBehaviour();
			},
		);
		this.#share = new ViewerShare(this.#root, this.controller, () => this.#profile.profile);
		this.#share.onChange = (detail) =>
			this.dispatchEvent(
				new CustomEvent('collaboration-change', { detail, bubbles: true, composed: true }),
			);
		this.#share.onError = (error) =>
			this.dispatchEvent(
				new CustomEvent('document-error', { detail: error, bubbles: true, composed: true }),
			);
		this.#backstage = new ViewerBackstage({
			root: this.#root,
			viewport: this.#viewport,
			fileName: () => this.#fileName,
			load: (source) => this.load(source),
			createBlankDrawing: () => this.createBlankDrawing(),
			createFromTemplate: (template) => this.createFromTemplate(template),
			exportVsdx: () => this.controller.exportVsdx(),
			// A downloaded copy exists as a file: the AutoRecover copy has done its job.
			downloaded: () => this.#recover.saved(),
			exportSvg: () => this.exportSvg(),
			exportPicture: async (format, choice) => {
				const { document: model, pageIndex } = this.controller.state;
				if (!model) throw new Error('Open a drawing before exporting.');
				const doc = this.ownerDocument;
				if (format === 'png')
					return { blob: await exportPagePng(doc, model, pageIndex), pageIndex };
				const result = await exportDocumentPdfDetailed(doc, model, {
					...pdfOptionsFor(choice ?? { scope: 'all' }, pageIndex),
					fonts: this.pdfFonts,
				});
				return { blob: result.blob, pageIndex, report: result };
			},
			closeDocument: () => {
				this.#fileName = '';
				this.controller.setDocument(null);
			},
			showOptions: () => this.#profile.showOptions(),
			announce: (message) => {
				this.#announcement = message;
				this.#status.textContent = message;
			},
			noteCount: () => this.#notes.children.length,
			pageSetup: () =>
				this.#commands.run({ type: 'page-setup', command: { op: 'dialog', tab: 'print' } }),
			fileCommand: (command) => announceFileCommand(this, command),
			storage: () => this.#storage,
			revision: () => this.controller.state.document,
			savedToStorage: (name, revision) => {
				this.#fileName = name;
				const { state } = this.controller;
				// The host stored it: the drawing is saved, unless it was edited while the save ran.
				// Then it stays dirty and keeps its AutoRecover copy, which may hold the newer edits.
				if (state.document === revision && !state.edit.busy && !state.loading) {
					this.controller.markClean();
					this.#recover.saved();
				}
				this.#render(this.controller.state);
			},
			saveFailed: (error) =>
				this.dispatchEvent(
					new CustomEvent('document-error', { detail: error, bubbles: true, composed: true }),
				),
			customization: () => this.#custom.resolved,
			regionContext: (region) => ({
				region,
				viewer: this,
				locale: this.#locale(),
				canEdit: this.controller.canEdit,
				pageIndex: this.controller.state.pageIndex,
				pageCount: this.controller.state.document?.pages.length ?? 0,
				close: () => this.#backstage.hide(),
			}),
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
	/**
	 * The status bar's look: `'neutral'` (the suite's bar, the default) or `'colorful'` (Visio's
	 * product colour). Reflects the `status-bar` attribute; `--vv-status-background` and
	 * `--vv-status-ink` recolour either.
	 */
	get statusBar(): StatusBarTheme {
		return this.getAttribute('status-bar') === 'colorful' ? 'colorful' : 'neutral';
	}
	set statusBar(value: StatusBarTheme) {
		this.setAttribute('status-bar', value === 'colorful' ? 'colorful' : 'neutral');
	}
	/** Whether the page the viewer sits in is using the dark theme. */
	#dark(): boolean {
		const root = this.ownerDocument.documentElement;
		return (root.dataset.officeTheme ?? root.dataset.theme) === 'dark';
	}
	/**
	 * Light or dark belongs to the page, which every editor on it shares. Ask the host with a
	 * cancelable `office-theme-request` (`{ scheme }`); when no host takes it, set the shared
	 * theme's `data-office-theme` on the document.
	 */
	#requestScheme(scheme: 'light' | 'dark'): void {
		const event = new CustomEvent('office-theme-request', {
			detail: { scheme },
			bubbles: true,
			composed: true,
			cancelable: true,
		});
		if (this.dispatchEvent(event)) this.ownerDocument.documentElement.dataset.officeTheme = scheme;
	}
	/**
	 * Tabs a host adds after Help, as an Office add-in does (Visio's ACROBAT tab). Each command
	 * runs its `run` callback and dispatches `office-ribbon-add-in` (`{ tab, command }`) from the
	 * viewer. Assign a new array to change them; a tab cannot take a built-in tab's id.
	 */
	get ribbonAddIns(): readonly RibbonAddInTab[] {
		return this.#ribbonAddIns;
	}
	set ribbonAddIns(value: readonly RibbonAddInTab[]) {
		this.#assertAlive();
		this.#ribbonAddIns = [...(value ?? [])];
		syncRibbonAddIns(this.#toolbar, this.#ribbonAddIns, {
			panelClass: 'ribbon-content',
			wrap: (doc, tab, groups) => [commandRow(doc, `${tab.label} commands`, groups)],
		});
	}
	get showToolbar(): boolean {
		return !this.#toolbar.hidden;
	}
	set showToolbar(value: boolean) {
		this.#assertAlive();
		this.#toolbar.hidden = !value;
		this.#titleBar.hidden = !value;
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
		// Now the drawing has its name: offer back a copy kept from an earlier, unsaved session.
		this.#recover.opened();
	}
	/** Create a source-backed drawing and name only the accepted request. */
	async createBlankDrawing(options?: CreateVsdxOptions): Promise<void> {
		return this.#createDrawing(() => this.controller.createBlankDrawing(options));
	}
	/** Create a drawing from a File > New template (Blank Drawing, flowcharts, org chart). */
	async createFromTemplate(template: VisioTemplateId): Promise<void> {
		return this.#createDrawing(() => this.controller.createFromTemplate(template));
	}
	async #createDrawing(create: () => Promise<void>): Promise<void> {
		this.#assertAlive();
		const token = ++this.#loadToken;
		const generation = this.controller.documentGeneration;
		await create();
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
		this.controller.cancelEdit();
	}
	/** Visio's status bar: Page n of m, then the selection's Width, Height and Angle, and language. */
	#renderStatus(state: ViewerState): void {
		const set = (name: string, value: string) =>
			this.#root.querySelector(`[data-${name}-status]`)!.setAttribute('value', value);
		const shape = selectedShape(state.document, state.selectedShape, state.pageIndex);
		const inches = (value: number) => `${+value.toFixed(3)} in.`;
		const many = state.selectedShapes.length > 1;
		set(
			'width',
			many
				? `${state.selectedShapes.length} shapes selected`
				: shape
					? `Width: ${inches(shape.width)}`
					: '',
		);
		set('height', !many && shape ? `Height: ${inches(shape.height)}` : '');
		// Visio's Angle is counter-clockwise, in (-180°, 180°].
		const turn = shape?.rotation
			? ((((shape.rotation.angle * 180) / Math.PI) % 360) + 360) % 360
			: 0;
		const angle = turn > 180 ? turn - 360 : turn;
		set('angle', !many && shape ? `Angle: ${+angle.toFixed(2)}°` : '');
		set(
			'language',
			state.document
				? statusLanguage(this.lang || this.ownerDocument.documentElement.lang || navigator.language)
				: '',
		);
	}
	/** Shape text edits in place on the canvas; compatibility notes live in File > Info. */
	#reveal(panel: 'edit' | 'notes' | 'selection' | 'format'): void {
		if (panel === 'edit') this.#edit.start();
		else if (panel === 'notes') this.#backstage.show('info');
		else this.#chrome.reveal(panel);
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
		this.#share.stop();
		this.#disposeInputs();
		this.#fontEvents?.removeEventListener('loadingdone', this.#fontsChanged);
		this.#fontEvents = undefined;
		this.#canvas.dispose();
		this.#unsubscribe();
		this.#eventUnsubscribe();
		this.controller.destroy();
		this.#root.replaceChildren();
	}
	attributeChangedCallback(): void {
		this.#localizeRibbon();
	}
	/** Put the File > Options behaviour settings to work on the pieces they govern. */
	#applyBehaviour(): void {
		const behaviour = this.#behaviour;
		this.controller.undoLevels = behaviour.undoLevels;
		this.#edit.spellcheck = behaviour.checkSpelling;
		this.#edit.autoSize = behaviour.autoSizeText;
		this.#mini.enabled = behaviour.miniToolbar;
		this.#livePreview.enabled = behaviour.livePreview;
		this.#recover.configure(behaviour.autoRecover, behaviour.autoRecoverMinutes);
		if (!behaviour.livePreview) this.#livePreview.stop();
		const ribbon = this.#toolbar as unknown as HTMLElement & {
			selected: string;
			setTabHidden(tab: string, reason: string, hidden: boolean): void;
		};
		const shown = CUSTOMIZABLE_TABS.filter(([id]) => behaviour[`tab-${id}`]).map(([id]) => id);
		for (const [id] of CUSTOMIZABLE_TABS) ribbon.setTabHidden?.(id, 'options', !shown.includes(id));
		// A hidden tab cannot stay the open one.
		if (shown.length && !(shown as readonly string[]).includes(ribbon.selected))
			ribbon.selected = shown[0]!;
		this.#localizeRibbon();
		// The title bar is built before the profile exists; it redraws with the first render.
		if (this.#profile)
			renderTitleBar(
				this.#titleBar,
				this.controller.state,
				this.#fileName,
				this.#profile.profile,
				quickAccessShown(behaviour, this.#custom.resolved),
			);
	}
	/**
	 * The ribbon speaks the `locale` attribute, else the language chosen in Options, else `lang`,
	 * else the page's, else the browser's. A host that sets `locale` itself keeps its choice.
	 */
	#locale(): string {
		return (
			this.getAttribute('locale') ||
			this.#behaviour.displayLanguage ||
			this.lang ||
			this.ownerDocument.documentElement.lang ||
			navigator.language
		);
	}
	#localizeRibbon(): void {
		if (this.#disposed || !this.#overflow) return;
		// Translation rewrites tooltips from their full text, so the style is applied afterwards.
		restoreScreenTips(this.#toolbar);
		localizeRibbon(this.#toolbar, this.#locale());
		applyScreenTips(this.#toolbar, this.#behaviour.screenTips);
		this.#overflow.refit();
	}
	connectedCallback(): void {
		if (this.#disposed) return;
		installOfficeUiThemeFor(this);
		// Office Theme: the host's status-bar attribute wins; otherwise the choice saved in Options.
		const savedTheme = readStatusBarTheme();
		if (!this.hasAttribute('status-bar') && savedTheme) this.statusBar = savedTheme;
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
		const disposeControls = this.#controlHandles.wire();
		const disposeRotation = this.#rotationHandle.wire();
		const disposeResize = this.#resizeHandles.wire();
		const disposeBackstage = this.#backstage.wire();
		const disposeMini = this.#mini.wire(this.#viewport);
		const disposePreview = this.#livePreview.wire();
		const disposeRecover = this.#recover.wire(this.ownerDocument);
		const disposeMenus = wireContextMenus(
			this.#root,
			this.#viewport,
			this.controller,
			(x, y) => this.#mini.showAt(x, y),
			(menu, target) => {
				syncVisioHostMenu(menu, target, this.#custom.resolved, this.controller);
				syncShapeActions(menu, target, this.controller, (message) => {
					this.#announcement = message;
					this.#status.textContent = message;
				});
			},
		);
		const disposeTellMe = wireTellMe(this.#root);
		const disposeGalleries = wireInlineGalleries(this, this.#root);
		const keyTips = attachKeyTips(this.#root);
		const disposePanZoom = this.#panZoom.wire();
		const disposeSizePosition = this.#sizePosition.wire();
		const disposeInlineSize = this.#inlineSize.wire();
		const disposeProfile = this.#profile.wire();
		// The title bar's Quick Access Toolbar: Save downloads a copy, Undo and Redo use the history.
		const quickAccess = (event: Event): void => {
			const id = (event as CustomEvent<{ command: string }>).detail.command;
			event.stopPropagation();
			if (id === 'save') this.#backstage.download();
			else if (id === 'undo' || id === 'redo') this.#commands.run({ type: 'history', key: id });
		};
		this.#titleBar.addEventListener('office-command', quickAccess);
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
		// Narrow windows collapse ribbon groups into buttons; add-in tabs are refitted as they change.
		const overflow = wireRibbonOverflow(this.#toolbar);
		const disposeTabCue = wireRibbonTabCue(this.#toolbar);
		this.#overflow = overflow;
		this.#localizeRibbon();
		const disposeReplace = this.#replace.wire();
		const disposeEdit = this.#edit.wire();
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
			(character) => this.#commands.tool === 'pointer' && this.#edit.start(character),
		);
		return () => {
			disposeChrome();
			disposeCommands();
			disposePresentation();
			disposePointer();
			disposeLineEndpoints();
			disposeControls();
			disposeRotation();
			disposeResize();
			disposeBackstage();
			disposeMenus();
			disposeMini();
			this.#recover.dispose();
			disposePreview();
			disposeRecover();
			disposeTellMe();
			disposeGalleries();
			keyTips.dispose();
			disposePanZoom();
			disposeSizePosition();
			disposeInlineSize();
			disposeProfile();
			this.#titleBar.removeEventListener('office-command', quickAccess);
			disposeShare();
			disposeRulers();
			disposeStencil();
			disposeFind();
			overflow.destroy();
			disposeTabCue();
			this.#overflow = undefined;
			disposeReplace();
			disposeInputs();
			disposeEdit();
		};
	}
	#assertAlive(): void {
		if (this.#disposed) throw new Error('The viewer has been destroyed.');
	}
	#render(state: ViewerState): void {
		if (this.#suspended) return;
		const page = state.document?.pages[state.pageIndex];
		// A preview belongs to the drawing it was made from; any real change ends it.
		if (this.#preview && this.#preview.base !== state.document) this.#preview = undefined;
		const changed = this.#canvas.render(
			this.#preview ? { ...state, document: this.#preview.document } : state,
		);
		const previous = this.#handleState;
		this.#handleState = state;
		if (
			changed ||
			!previous ||
			previous.canEdit !== state.canEdit ||
			previous.zoom !== state.zoom ||
			previous.selectedShape !== state.selectedShape ||
			previous.selectedShapes !== state.selectedShapes ||
			previous.loading !== state.loading ||
			previous.edit.busy !== state.edit.busy ||
			previous.edit.sourceAvailable !== state.edit.sourceAvailable
		) {
			this.#lineEndpoints.render(state);
			this.#controlHandles.render(state);
			this.#rotationHandle.render(state);
			this.#resizeHandles.render(state);
			this.#viewport.toggleAttribute(
				'data-selection-frame',
				!!this.#viewport.querySelector('[data-resize-overlay], [data-line-endpoint]'),
			);
		}
		if (changed) this.#announcement = undefined;
		this.#zoomSlider.value = Math.round(state.zoom * 100);
		this.#zoomSlider.disabled = !page;
		this.#renderStatus(state);
		renderTitleBar(
			this.#titleBar,
			state,
			this.#fileName,
			this.#profile.profile,
			quickAccessShown(this.#behaviour, this.#custom.resolved),
		);
		for (const button of this.#root.querySelectorAll<HTMLButtonElement>('[data-action]'))
			button.disabled = !page;
		renderFindBar(this.#findBar, state);
		this.#replace.render(state);
		this.#edit.render(state);
		this.#chrome.render(state);
		this.#commands.render(state);
		renderContextualTabs(this.#toolbar, state);
		if (this.#mini.visible) this.#mini.sync();
		this.#presentation.render(state);
		this.#backstage.render(state);
		this.#pointer.render(state);
		this.#panZoom.render();
		this.#sizePosition.render(state);
		this.#inlineSize.render(state);
		this.#viewport.setAttribute('aria-busy', String(state.loading || state.edit.busy));
		this.#status.textContent = state.loading
			? 'Opening diagram…'
			: state.edit.busy
				? 'Updating diagram…'
				: (state.error?.message ??
					(state.edit.error ? `Edit rejected: ${editErrorMessage(state.edit.error)}` : undefined) ??
					this.#announcement ??
					'');
		// Routine announcements are for assistive technology; errors also show, as Visio's do.
		this.#status.parentElement!.toggleAttribute(
			'data-error',
			!state.loading && !state.edit.busy && !!(state.error ?? state.edit.error),
		);
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
