import { syncRibbonAddIns } from '../ribbon/add-in-tabs';
import { refitRibbon } from './ribbon-overflow';
import { applyShellLabels } from './shell-labels';
import type { DocumentModel } from 'ooxml-core/docx';
import { createDocument } from 'ooxml-core/docx';
import type { EditorCore } from './editor-core';
import { emit, on } from './events';
import { createRibbon } from './ribbon';
import type { RibbonActionId } from 'ooxml-core/docx/ui';
import { editorStyleText } from './styles';
import { runRibbonCommand } from './editor-commands';
import { createSearchPanel } from './search-panel';
import { createPrintLayoutController } from './print-layout-view';
import { createPrintEditingBridge } from './print-editing';
import { attachPrintNoteEditing } from './print-note-editing';
import { attachPrintFloatSelection } from './print-float-selection';
import { attachPrintStoryEditing } from './print-story-editing';
import { moveCursorToBlock } from './print-layout-cursor';
import { ReviewController } from './review-controller';
import { EditorChrome } from './editor-chrome';
import { dispatchDocumentError } from './editor-host';
import { routeRibbonAction } from './ribbon-router';
import { attachSurfaceNavigation } from './surface-navigation';
import { documentWordCount } from 'ooxml-core/docx/ui';
import { PageNavigator } from './page-navigator';
import { HeadingNavigator } from './heading-navigator';
import { applyViewOptions } from './view-options';
import { wordRegionContext } from './customization-apply';
import { attachEditorInteractions } from './editor-interactions';
import { attachCanvasClick } from './canvas-click';
import { createChartFormatPane } from './chart-format-pane';

/** Element-level operations the chrome needs; they stay on the element because they are public API. */
export interface ShellApi {
	setReadOnly(readOnly: boolean): void;
	setDocumentModel(model: DocumentModel): void;
	load(bytes: Uint8Array | ArrayBuffer): Promise<void>;
	saveBytes(): Promise<Uint8Array>;
}

function createChrome(core: EditorCore, api: ShellApi): EditorChrome {
	const { element, shell } = core;
	const host = () => element as unknown as { hiddenActions: readonly RibbonActionId[] };
	return new EditorChrome({
		element,
		model: () => core.model,
		ribbon: () => shell.toolbar,
		wordCount: () => {
			const doc = core.view?.state.doc;
			return doc ? documentWordCount(doc, element.lang || undefined) : 0;
		},
		locale: () => core.locale,
		readOnly: () => core.readOnly,
		setReadOnly: (readOnly) => {
			core.reportReadOnly(() => api.setReadOnly(readOnly));
		},
		newDocument: () => api.setDocumentModel(createDocument()),
		load: (bytes) => api.load(bytes),
		save: () => api.saveBytes(),
		storage: () => core.storage,
		saveStateChanged: (state) => {
			if (state !== 'saving') core.dirtyState.set(state === 'dirty');
		},
		print: () => core.pages.print(),
		history: (key) => {
			if (!core.view) return;
			if (core.parts.usesBodyHistory()) core.parts.runHistory(key);
			else runRibbonCommand(core.view, { type: 'history', key }, core.collab.ids);
			if (typeof document.execCommand === 'function') core.targetView()?.focus();
		},
		toggleComments: () => shell.review?.handleComments('toggle'),
		collaborating: () => core.collab.active,
		setViewMode: (mode) => core.pages.setViewMode(mode),
		setZoom: (percent) => core.pages.setZoom(percent),
		options: () => ({
			locale: core.locale,
			theme: core.theme,
			author: core.reviewAuthor,
			spellCheck: core.view?.dom.spellcheck !== false,
			showRuler: Boolean(shell.canvas?.querySelector('.dve-ruler')),
			showMarks: shell.paper?.hasAttribute('data-show-marks') ?? false,
			showThumbnails: core.viewOptions.showThumbnails,
		}),
		setOption: (key, value) => {
			if (key === 'spellCheck') {
				if (core.view) core.view.dom.spellcheck = value === 'true';
				core.refreshControls();
			} else if (key === 'showThumbnails')
				element.toggleAttribute('show-thumbnails', value === 'true');
			else if (key === 'showRuler' || key === 'showMarks') {
				const current =
					key === 'showRuler'
						? Boolean(shell.canvas?.querySelector('.dve-ruler'))
						: (shell.paper?.hasAttribute('data-show-marks') ?? false);
				if (current !== (value === 'true'))
					core.inserts.handle({ type: key === 'showRuler' ? 'ruler' : 'showMarks' });
			} else element.setAttribute(key === 'author' ? 'review-author' : key, value);
		},
		hiddenActions: () => host().hiddenActions,
		setHiddenActions: (ids) => {
			host().hiddenActions = ids as RibbonActionId[];
			emit(element, 'ribbon-customize', host().hiddenActions);
		},
		customization: () => core.viewOptions.customization,
		regionContext: (region) => wordRegionContext(core, region),
		reportError: (error) => dispatchDocumentError(element, error),
	});
}

/** Builds the shadow-DOM shell (ribbon, canvas, panels, chrome) once; later calls are no-ops. */
export function buildShell(core: EditorCore, api: ShellApi): void {
	const { element, shell } = core;
	if (shell.paper) {
		attachPrintEditing(core);
		if (core.surfaceOnly) attachSurface(core);
		return;
	}
	element.classList.add('dve-host');
	const root = element.attachShadow({ mode: 'open' });
	const style = document.createElement('style');
	style.textContent = editorStyleText;
	const frame = document.createElement('section');
	frame.className = 'dve-frame';
	const toolbar = core.surfaceOnly ? undefined : createRibbon(core.locale);
	if (toolbar) on(toolbar, 'ribbon-action', (event) => routeRibbonAction(core, event.detail));
	const canvas = document.createElement('main');
	canvas.className = 'dve-canvas';
	const paper = document.createElement('div');
	paper.className = 'dve-paper';
	canvas.append(paper);
	const printLayout = createPrintLayoutController(
		canvas,
		(blockId, offset) => {
			core.pages.setViewMode('draft');
			if (core.view) moveCursorToBlock(core.view, blockId, offset);
		},
		(partName, contentType) => core.imageMedia.urlFor(partName, contentType),
		() => core.refreshControls(),
		() => core.reviewDisplayMode,
		() => core.pages.viewMode === 'print',
	);

	canvas.append(printLayout.element);
	let scrollFrame: number | undefined;
	canvas.addEventListener(
		'scroll',
		() => {
			if (core.pages.viewMode !== 'print' || scrollFrame !== undefined) return;
			scrollFrame = requestAnimationFrame(() => {
				scrollFrame = undefined;
				if (!element.isConnected || core.pages.viewMode !== 'print') return;
				core.refreshPageStatus();
			});
		},
		{ passive: true },
	);
	shell.printLayout = printLayout;
	if (core.surfaceOnly) {
		frame.append(canvas);
		root.append(style, frame);
		Object.assign(shell, { canvas, paper });
		attachPrintEditing(core);
		attachCanvasClick(core);
		attachSurface(core);
		applyShellLabels(element, paper, core.locale);
		return;
	}
	if (!toolbar) return;
	shell.navigator = new PageNavigator({
		pages: () => printLayout.pageElements(),
		pageSnapshot: (index) => printLayout.pageSnapshot(index),
		version: () => printLayout.layoutVersion(),
		goTo: (page) => {
			printLayout.scrollToPage(page);
			core.refreshControls();
		},
		close: () => element.toggleAttribute('show-thumbnails', false),
	});
	shell.searchPanel = createSearchPanel({
		getView: () => core.view,
		onClose: () => core.view?.focus(),
	});
	shell.headings = new HeadingNavigator({
		model: () => core.model,
		view: () => core.view,
		goTo: (id) => {
			core.pages.setViewMode('draft');
			if (core.view) moveCursorToBlock(core.view, id, 0);
			core.view?.dispatch(core.view.state.tr.scrollIntoView());
		},
		close: () => {
			shell.headings?.setOpen(false);
			core.refreshControls();
			core.view?.focus();
		},
	});
	shell.searchPanel.setLocale(core.locale);
	const review = new ReviewController({
		getModel: () => core.model,
		setModel: (model) => {
			core.model = model;
		},
		getView: () => core.view,
		getReviewAuthor: () => core.reviewAuthor,
		getCollaborationIds: () => core.collab.ids,
		getYjs: () => core.collab.yjs,
		canEditComments: () =>
			core.canEditBody() && (!core.collab.active || Boolean(core.collab.yjs?.sharedComments)),
		notifyChange: () => core.notifyChange(),
		refresh: () => core.refreshControls(),
	});
	shell.review = review;
	review.setLocale(core.locale);
	core.inserts.setLocale(core.locale);
	core.formatDialogs.setLocale(core.locale);
	const body = document.createElement('div');
	body.className = 'dve-body';
	shell.chartFormat = createChartFormatPane({
		view: () => core.targetView(),
		theme: () => core.model.theme,
		locale: () => core.locale,
		canEdit: () => core.canEditBody(),
	});
	body.append(
		shell.navigator.element,
		shell.headings.element,
		canvas,
		core.inserts.smartArt.pane,
		review.commentsPanel.element,
		shell.chartFormat.element,
		core.formatDialogs.restrictEditingPane,
	);
	frame.append(toolbar, shell.searchPanel.element, body);
	const chrome = createChrome(core, api);
	shell.chrome = chrome;
	chrome.mount(frame, toolbar);
	frame.append(
		core.inserts.linkDialog.element,
		core.inserts.pictureDialog.element,
		core.inserts.smartArt.dialog.element,
		...core.formatDialogs.elements.filter((el) => el !== core.formatDialogs.restrictEditingPane),
		core.inserts.pictureInput,
	);
	if (core.pendingFileName) chrome.fileName = core.pendingFileName;
	root.append(style, frame);
	Object.assign(shell, { toolbar, canvas, paper });
	attachPrintEditing(core);
	applyRibbonAddIns(core);
	// After the toolbar is registered: the File tab label is localized through the shell.
	chrome.setLocale(core.locale);
	applyViewOptions(core);
	chrome.applyCustomization();
	attachEditorInteractions(core, frame);
	element.setAttribute('role', 'region');
	applyShellLabels(element, paper, core.locale);
}

/** Renders the host's add-in tabs (`ribbonAddIns`) after Word's own, once the ribbon exists. */
export function applyRibbonAddIns(core: EditorCore): void {
	const toolbar = core.shell.toolbar;
	if (!toolbar) return;
	const before = toolbar.querySelector('[data-add-in]');
	syncRibbonAddIns(toolbar, core.ribbonAddIns, { panelClass: 'ribbon-panel' });
	// Equal tabs keep their panels; only new panels change what fits.
	if (toolbar.querySelector('[data-add-in]') !== before) refitRibbon(toolbar);
}

/** Read-only navigation for the surface; reconnection binds it again. */
function attachSurface(core: EditorCore): void {
	core.shell.surfaceNavigation ??= attachSurfaceNavigation(core);
}

/** A disconnected editor releases input listeners; reconnection binds the retained shell again. */
function attachPrintEditing(core: EditorCore): void {
	const { shell } = core;
	if (shell.printEditing || !shell.printLayout || !shell.paper) return;
	// Before the bridge, so a selected picture takes the arrow keys ahead of caret navigation.
	shell.printFloats = attachPrintFloatSelection({
		pages: shell.printLayout.element,
		paper: shell.paper,
		view: () => core.view,
		printActive: () => core.pages.viewMode === 'print',
	});
	shell.printEditing = createPrintEditingBridge(
		shell.printLayout.element,
		shell.paper,
		() => core.view,
		shell.printLayout.resolveClick,
		shell.canvas,
		(id, offset) => shell.printLayout?.ensureBlock(id, offset),
	);
	shell.printEditing.setActive(core.pages.viewMode === 'print');
	const { printLayout } = shell;
	shell.printStoryEditing = attachPrintStoryEditing({
		pages: printLayout.element,
		parts: core.parts,
		model: () => core.model,
		printActive: () => core.pages.viewMode === 'print',
		holdLayout: (held, page) => printLayout.holdLayout(held, page),
	});
	shell.printNoteEditing = attachPrintNoteEditing({
		pages: printLayout.element,
		parts: core.parts,
		model: () => core.model,
		view: () => core.view,
		printActive: () => core.pages.viewMode === 'print',
		holdLayout: (held, page) => printLayout.holdLayout(held, page),
		showBlock: (id, offset) => printLayout.ensureBlock(id, offset),
	});
}
