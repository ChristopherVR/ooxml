import type { DocumentStorage } from 'ooxml-core/embedding';
import type { SurfaceHostActions } from './building-blocks-actions';
import type { RibbonAddInTab } from '../ribbon/add-in-tabs';
import { FormatDialogs } from './format-dialogs';
import type { EditorState, Transaction } from 'prosemirror-state';
import type { EditorView } from 'prosemirror-view';
import type { DocumentModel } from 'ooxml-core/docx';
import { createDocument } from 'ooxml-core/docx';
import type { loadDocument } from 'ooxml-core/docx/load';
import { refreshEditorControls } from './editor-controls';
import { assignMissingParagraphIds, docToModel } from './model-adapter';
import type { SearchPanelHandle } from './search-panel';
import { CollaborationSession } from './collaboration-session';
import { roomFor, type DocxRoom } from './collaboration-room';
import { repairCollaborativeDocumentIds, isWordYjsRemoteTransaction } from 'ooxml-core/docx/ui';
import type { EditorTheme, EditorThemeMode } from 'ooxml-core/docx/ui';
import type { EditorLocale } from './localization';
import { findLocalizedControl, formatPageStatus } from './localization';
import type { PrintEditingBridge } from './print-editing';
import type { PrintFloats } from './print-float-selection';
import type { PrintLayoutController } from './print-layout-view';
import { REMOTE_TRANSACTION_META } from './track-changes-mode';
import type { ReviewDisplayMode } from './review-display';
import type { ReviewController } from './review-controller';
import { ImageMediaCache } from './image-media';
import type { EditorChrome } from './editor-chrome';
import { dispatchDocumentError, type EditorHost } from './editor-host';
import { PartsController } from './parts-controller';
import { PageController } from './page-controller';
import { InsertController } from './insert-controller';
import { emit } from './events';
import { reflectAttribute } from './editor-attributes';
import { DirtyState } from 'ooxml-core/docx/ui';
import { PageTracker, syncPageState } from './page-sync';
import type { PageNavigator } from './page-navigator';
import type { HeadingNavigator } from './heading-navigator';
import type { DocxChartFormatPane } from './chart-format-pane';
import { HEADER_FOOTER_INPUT, locksStories } from 'ooxml-core/docx/ui';
import { newPicturePartName } from './picture-commands';
import { ViewOptions } from './view-options';
import { currentSectionIndex } from './section-commands';
import { syncHeaderFooterRibbon } from './header-footer-ribbon';
import { syncSmartArtDesign } from './smartart-design-ribbon';

export type LoadedDocument = Awaited<ReturnType<typeof loadDocument>>;

/** DOM pieces created by the shell once the element connects. */
export interface ShellParts {
	toolbar?: HTMLElement;
	canvas?: HTMLElement;
	paper?: HTMLElement;
	searchPanel?: SearchPanelHandle;
	review?: ReviewController;
	chrome?: EditorChrome;
	printLayout?: PrintLayoutController;
	printEditing?: PrintEditingBridge;
	/** Detaches in-page header and footer editing. */
	printStoryEditing?: () => void;
	/** Detaches in-page footnote and endnote editing. */
	printNoteEditing?: () => void;
	/** Selection frame and gestures for floating pictures on the pages. */
	printFloats?: PrintFloats;
	/** Detaches read-only selection and keyboard navigation of the surface. */
	surfaceNavigation?: () => void;
	navigator?: PageNavigator;
	headings?: HeadingNavigator;
	/** The docked Format pane of the selected chart. */
	chartFormat?: DocxChartFormatPane;
}

/**
 * Mutable editor state and the controllers that share it. `<docx-editor>` is a thin facade over
 * this object; the shell, render, ribbon and chrome modules take it instead of reaching into the
 * element's private fields.
 */
export class EditorCore {
	model: DocumentModel = createDocument();
	/** Tabs the host added after Word's own (`ribbonAddIns` of the element). */
	ribbonAddIns: readonly RibbonAddInTab[] = [];
	view?: EditorView | undefined;
	loaded?: LoadedDocument | undefined;
	loadGeneration = 0;
	loading = false;
	/** The host's `readOnly` switch. A collaboration `viewer` role locks the editor on top of it. */
	hostReadOnly = false;
	/** True while `applyReadOnly` writes the `read-only` attribute. */
	reflectingReadOnly = false;
	private shownReadOnly = false;
	private shownHostReadOnly = false;
	/**
	 * Effective: the host's switch, a `viewer` role, or `capabilities.edit` off. Assigning sets
	 * the host's switch only, so nothing that assigns it can lift the other two.
	 */
	get readOnly(): boolean {
		return this.hostReadOnly || this.roleLocked || !this.viewOptions.customization.canEdit;
	}
	set readOnly(value: boolean) {
		this.hostReadOnly = value;
	}
	/** Applies the effective value to the view, controls and attribute. */
	applyReadOnly(): void {
		this.view?.setProps({ editable: () => this.canEditBody() });
		if (this.readOnly) this.parts.closeHeaderFooter(this.shell.canvas, this.shell.paper);
		this.refreshControls();
		// Reflecting the attribute re-enters attributeChangedCallback; that echo must not apply again.
		this.reflectingReadOnly = true;
		try {
			reflectAttribute(this.element, 'read-only', this.hostReadOnly);
		} finally {
			this.reflectingReadOnly = false;
		}
		this.shownReadOnly = this.readOnly;
		this.shownHostReadOnly = this.hostReadOnly;
	}
	/** True while a collaboration `viewer` role holds the editor read-only. */
	get roleLocked(): boolean {
		return this.collab.yjs?.session.role === 'viewer';
	}
	/** Applies the effective value, then reports it when it or the host's switch changed. */
	reportReadOnly(change?: () => void): void {
		const wasReadOnly = this.shownReadOnly;
		const wasHost = this.shownHostReadOnly;
		change?.();
		this.applyReadOnly();
		const readOnly = this.readOnly;
		if (readOnly === wasReadOnly && this.hostReadOnly === wasHost) return;
		emit(this.element, 'readonly-change', {
			readOnly,
			locked: this.roleLocked,
			hostReadOnly: this.hostReadOnly,
		});
	}
	/** Joining or leaving a room can change the role lock: apply it and report a change. */
	roleChanged(): void {
		this.reportReadOnly();
	}
	/** Set before mounting to render only the document surface. */
	surfaceOnly = false;
	/** Host-owned workflows for surface mode (link, find and replace, picture picker). */
	surfaceActions: SurfaceHostActions = {};
	/** Where File > Save and Open go instead of a download and the file picker (`storage`). */
	storage: DocumentStorage | undefined = undefined;
	locale: EditorLocale = 'en';
	reviewAuthor = 'Author';
	theme: EditorThemeMode = 'auto';
	themeColors?: Partial<EditorTheme> | undefined;
	appliedThemeVars: string[] = [];
	reviewDisplayMode: ReviewDisplayMode = 'all';
	detachedState?: EditorState | undefined;
	pendingFileName?: string;
	readonly shell: ShellParts = {};
	readonly collab: CollaborationSession;
	/** The declarative `collaboration` property and the Yjs state reporting. */
	readonly room: DocxRoom;
	readonly inserts: InsertController;
	readonly formatDialogs: FormatDialogs;
	readonly imageMedia: ImageMediaCache;
	readonly host: EditorHost;
	readonly parts: PartsController;
	readonly pages: PageController;
	readonly viewOptions = new ViewOptions();
	readonly dirtyState: DirtyState;
	readonly pageTracker: PageTracker;
	private wordText = '';

	constructor(readonly element: HTMLElement) {
		this.dirtyState = new DirtyState((dirty) => emit(element, 'dirty-change', dirty));
		this.pageTracker = new PageTracker(element);
		this.collab = new CollaborationSession(element, () => this.view);
		this.room = roomFor(this);
		this.collab.onYjsChange = (session) => this.room.yjsChanged(session);
		this.inserts = new InsertController({
			view: () => this.targetView(),
			model: () => this.model,
			contentWidth: () => this.contentWidth(),
			paper: () => this.shell.paper,
			toolbar: () => this.shell.toolbar,
			reportError: (error) => emit(element, 'document-error', error),
			refresh: () => this.refreshControls(),
			newMediaPartName: (contentType) =>
				this.collab.yjs?.media.partName(contentType) ?? newPicturePartName(contentType),
			stageMedia: (name, part) => this.collab.yjs?.media.publish(name, part),
		});
		this.formatDialogs = new FormatDialogs({
			view: () => this.targetView(),
			reviewView: () => this.view,
			historyView: () => (this.parts.usesBodyHistory() ? this.view : this.targetView()),
			model: () => this.model,
			canDefineList: () => !this.readOnly && !this.collab.active,
			theme: () => this.model.theme,
			zoom: {
				percent: () => Math.round(this.pages.zoom * 100),
				setPercent: (percent) => this.pages.setZoom(percent),
				fit: (mode) => this.pages.zoomTo(mode),
			},
			pageSetup: {
				section: () => this.pages.currentSection(),
				canEdit: () => !this.readOnly && !this.collab.active,
				apply: (values) => this.pages.applyPageSetupValues(values),
			},
			lineNumbers: {
				section: () => this.pages.currentSection(),
				canEdit: () => !this.readOnly && !this.collab.active,
				apply: (settings) => this.pages.applyLineNumberSettings(settings),
			},
			watermark: {
				current: () => this.pages.currentWatermark(),
				canEdit: () => !this.readOnly && !this.collab.active,
				apply: (spec) => this.pages.applyWatermark(spec),
			},
			pageBorders: {
				section: () => this.pages.currentSection(),
				canEdit: () => !this.readOnly && !this.collab.active,
				apply: (borders) => this.pages.applyPageBorders(borders),
			},
			columns: {
				section: () => this.pages.currentSection(),
				canEdit: () => !this.readOnly && !this.collab.active,
				apply: (columns, footnoteColumns) => this.pages.applyColumns(columns, footnoteColumns),
			},
			restrictEditing: {
				protection: () => this.model.protection,
				apply: (protection) => this.applyProtection(protection),
			},
		});
		this.imageMedia = new ImageMediaCache(
			(partName) =>
				this.collab.yjs?.media.get(partName)?.bytes ??
				this.inserts.media(partName, this.loaded?.media),
		);
		this.host = {
			element,
			view: () => this.view,
			model: () => this.model,
			setModel: (model) => {
				this.model = model;
			},
			locale: () => this.locale,
			// A document protected for forms lets only its body's fields change, as in Word.
			canEditOutsideBody: () =>
				!this.readOnly && !this.collab.active && !locksStories(this.model.protection),
			edited: () => this.markEditedOutsideBody(),
			reportError: (cause) => dispatchDocumentError(element, cause),
		};
		this.parts = new PartsController({
			...this.host,
			reviewDisplayMode: () => this.reviewDisplayMode,
			images: () => this.imageMedia,
			plugins: () => this.inserts.plugins(),
			sectionIndex: () => (this.view ? currentSectionIndex(this.view, this.model) : 0),
			refreshControls: () => this.refreshControls(),
			keepOpenWithin: () => [
				this.shell.toolbar,
				this.inserts.linkDialog.element,
				this.inserts.pictureDialog.element,
				this.inserts.smartArt.dialog.element,
				this.inserts.smartArt.pane,
				...this.formatDialogs.elements,
			],
		});
		this.pages = new PageController({
			...this.host,
			paper: () => this.shell.paper,
			toolbar: () => this.shell.toolbar,
			statusBar: () => this.shell.chrome?.statusBar,
			printLayout: () => this.shell.printLayout,
			printEditing: (active) => this.shell.printEditing?.setActive(active),
			refreshControls: () => this.refreshControls(),
		});
	}

	contentWidth(): number {
		return this.model.page.width - this.model.page.marginLeft - this.model.page.marginRight;
	}

	/** Starts or stops Review > Restrict Editing; saving writes it to settings.xml. */
	applyProtection(protection: DocumentModel['protection']): void {
		const { protection: _previous, ...rest } = this.model;
		this.model = protection ? { ...rest, protection } : rest;
		if (locksStories(protection) && this.parts.activeView())
			this.parts.closeHeaderFooter(this.shell.canvas, this.shell.paper);
		// Redraw the editable regions, and the ribbon's disabled commands.
		this.view?.dispatch(this.view.state.tr.setMeta('docx-protection', true));
		this.markEditedOutsideBody();
		this.refreshControls();
	}

	canEditBody(): boolean {
		return !this.readOnly && (!this.collab.yjs || this.collab.yjs.session.canWrite());
	}

	/** The ribbon acts on an open header/footer/note editor, otherwise on the main text. */
	targetView(): EditorView | undefined {
		return this.parts.activeView() ?? this.view;
	}

	notifyChange(): void {
		this.dirtyState.set(true);
		emit(this.element, 'document-change', this.model);
	}

	markEditedOutsideBody(): void {
		this.shell.chrome?.setSaveState('dirty');
		this.pages.relayout();
		this.notifyChange();
	}

	scheduleCollaborationSend(): void {
		this.collab.scheduleSend(() => this.pendingCollaboration());
	}

	pendingCollaboration() {
		const state = this.view?.state ?? this.detachedState;
		return state && this.collab.client ? this.collab.client.createPendingBatch(state) : null;
	}

	applyTransaction(transaction: Transaction, remote = false): void {
		const view = this.view;
		if (!view || view.isDestroyed) return;
		remote ||= Boolean(this.collab.yjs && isWordYjsRemoteTransaction(transaction));
		if (this.collab.yjs && this.readOnly && transaction.docChanged && !remote) {
			view.updateState(view.state);
			return;
		}
		if (remote) transaction.setMeta(REMOTE_TRANSACTION_META, true);
		const previousDoc = view.state.doc;
		const previousParts = view.state.doc.attrs.sectionParts;
		const previousNotes = view.state.doc.attrs.noteParts;
		const previousComments = view.state.doc.attrs.commentThreads;
		const result = view.state.applyTransaction(transaction);
		const applied = result.state;
		const docChanged = !applied.doc.eq(view.state.doc);
		remote ||= Boolean(
			this.collab.yjs &&
			!transaction.docChanged &&
			result.transactions.some(isWordYjsRemoteTransaction),
		);
		if (applied === view.state) {
			view.updateState(applied);
			return;
		}
		const repaired =
			remote || !docChanged
				? null
				: this.collab.active
					? repairCollaborativeDocumentIds(
							applied,
							this.collab.client?.clientId ?? String(this.collab.yjs!.session.clientId),
							this.collab.ids,
						)
					: assignMissingParagraphIds(applied);
		view.updateState(repaired ? applied.apply(repaired) : applied);
		if (docChanged) {
			this.model = docToModel(view.state.doc, this.model, previousDoc);
			if (previousComments !== view.state.doc.attrs.commentThreads)
				this.shell.review?.commentsPanel.refresh();
			this.pages.refreshPageStyles();
			this.pages.relayout();
			this.shell.chrome?.setSaveState('dirty');
			this.notifyChange();
			if (
				(previousParts !== view.state.doc.attrs.sectionParts ||
					previousNotes !== view.state.doc.attrs.noteParts) &&
				!transaction.getMeta(HEADER_FOOTER_INPUT)
			)
				this.parts.render(this.shell.canvas, this.shell.paper, true);
		}
		this.refreshControls();
		if (docChanged || remote) this.scheduleCollaborationSend();
		if (docChanged || transaction.selectionSet || remote) this.collab.presence?.schedule();
	}

	refreshControls(): void {
		this.element.dispatchEvent(new CustomEvent('editor-state-change'));
		this.parts.syncSection(this.shell.canvas, this.shell.paper);
		this.shell.headings?.sync(this.locale);
		if (this.shell.toolbar)
			findLocalizedControl(this.shell.toolbar, 'Navigation pane')?.setAttribute(
				'aria-pressed',
				String(this.shell.headings?.isOpen ?? false),
			);
		const { searchPanel, printLayout, review, chrome, toolbar } = this.shell;
		searchPanel?.refresh();
		this.shell.chartFormat?.refresh();
		// The open pane follows the caret's thread and whether a selection can take a comment.
		if (review?.commentsOpen) review.commentsPanel.refresh();
		if (this.pages.viewMode === 'print') {
			printLayout?.refreshCurrentPage();
			this.shell.printEditing?.refresh();
			this.shell.printFloats?.refresh();
		}
		const status = refreshEditorControls(
			toolbar,
			this.targetView(),
			this.model,
			!this.canEditBody(),
			this.collab.active,
			this.locale,
			this.element.lang,
			printLayout?.pageStatus(),
			Boolean(review?.commentsOpen),
			this.canEditBody() && (!this.collab.active || Boolean(this.collab.yjs?.sharedComments)),
		);
		this.inserts.smartArt.sync();
		const target = this.targetView();
		if (toolbar && target)
			syncSmartArtDesign(
				toolbar,
				target.state,
				target.editable && this.canEditBody(),
				this.inserts.smartArt.textPaneOpen,
			);
		if (status) {
			this.wordText = status.wordText;
			chrome?.refresh(status.pageText, status.wordText);
		}
		this.pages.syncControls();
		if (toolbar)
			syncHeaderFooterRibbon(
				toolbar,
				this.parts.headerFooterContext(),
				this.model,
				this.host.canEditOutsideBody(),
			);
		syncPageState(this);
		chrome?.ribbonActions.setCommentsOpen(Boolean(review?.commentsOpen));
	}

	/** Scrolling only changes page state, never formatting controls or document statistics. */
	refreshPageStatus(): void {
		const { printLayout, printEditing, chrome } = this.shell;
		printLayout?.refreshCurrentPage();
		printEditing?.refresh();
		const status = printLayout?.pageStatus();
		if (status)
			chrome?.statusBar.setPageAndWords(
				formatPageStatus(this.locale, status.current, status.total),
				this.wordText,
			);
		syncPageState(this);
		// A custom page control follows scrolling; the full editor has no listener for it.
		if (this.surfaceOnly) this.element.dispatchEvent(new CustomEvent('editor-state-change'));
	}
}
