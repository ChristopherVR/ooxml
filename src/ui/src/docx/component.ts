import { installOfficeUiThemeFor } from '../theme';
import { applyShellLabels } from './shell-labels';
import { EditorState } from 'prosemirror-state';
import { refreshWordYjsPresence } from 'ooxml-core/docx/ui';
import type { DocumentModel } from 'ooxml-core/docx';
import { createDocument } from 'ooxml-core/docx';
import { loadDocument } from 'ooxml-core/docx/load';
import { setRibbonLocale } from './ribbon';
import type { CollaborationConfig, ClientReceiveResult, StepBatch } from 'ooxml-core/docx/ui';
import { normalizeEditorLocale } from './localization';
import {
	applyThemeColors,
	normalizeThemeMode,
	type EditorTheme,
	type EditorThemeMode,
} from 'ooxml-core/docx/ui';
import type { DocxCollaborationOptions } from './collaboration-types';
import { leaveYjs } from './collaboration-room';
import { EditorCore, type LoadedDocument } from './editor-core';
import { buildShell, type ShellApi } from './editor-shell';
import { renderDocument } from './editor-render';
import { registerSmartArt } from './smartart-node-view';
import { DocxEditorApi } from './element-api';
import { emit, type DocxEditorEventMap } from './events';
import {
	DEFAULT_FILE_NAME,
	DEFAULT_REVIEW_AUTHOR,
	DOCX_EDITOR_ATTRIBUTES,
	applyAttribute,
	isDocxEditorAttribute,
	reflectAttribute,
} from './editor-attributes';

/**
 * Typed event overloads. Redeclaring the methods hides the inherited ones, so the standard DOM
 * event map and the string fallback are repeated here.
 */
export interface DocxEditorElement {
	addEventListener<K extends keyof DocxEditorEventMap>(
		type: K,
		listener: (this: DocxEditorElement, event: DocxEditorEventMap[K]) => unknown,
		options?: boolean | AddEventListenerOptions,
	): void;
	addEventListener<K extends keyof HTMLElementEventMap>(
		type: K,
		listener: (this: DocxEditorElement, event: HTMLElementEventMap[K]) => unknown,
		options?: boolean | AddEventListenerOptions,
	): void;
	addEventListener(
		type: string,
		listener: EventListenerOrEventListenerObject,
		options?: boolean | AddEventListenerOptions,
	): void;
	removeEventListener<K extends keyof DocxEditorEventMap>(
		type: K,
		listener: (this: DocxEditorElement, event: DocxEditorEventMap[K]) => unknown,
		options?: boolean | EventListenerOptions,
	): void;
	removeEventListener<K extends keyof HTMLElementEventMap>(
		type: K,
		listener: (this: DocxEditorElement, event: HTMLElementEventMap[K]) => unknown,
		options?: boolean | EventListenerOptions,
	): void;
	removeEventListener(
		type: string,
		listener: EventListenerOrEventListenerObject,
		options?: boolean | EventListenerOptions,
	): void;
}

export class DocxEditorElement extends DocxEditorApi {
	/** Attributes mirrored to the `locale`, `readOnly`, `fileName`, `reviewAuthor` and `theme` properties. */
	static get observedAttributes(): string[] {
		return [...DOCX_EDITOR_ATTRIBUTES];
	}

	protected readonly core = new EditorCore(this);
	private readonly shellApi: ShellApi = {
		setReadOnly: (readOnly) => {
			this.readOnly = readOnly;
		},
		setDocumentModel: (model) => {
			this.documentModel = model;
		},
		load: (bytes) => this.load(bytes),
		saveBytes: () => this.saveBytes(),
	};

	/** Internal seams that in-repo tests reach for; not public API. */
	private get view() {
		return this.core.view;
	}
	private get inserts() {
		return this.core.inserts;
	}

	attributeChangedCallback(name: string, _old: string | null, value: string | null) {
		if (isDocxEditorAttribute(name)) applyAttribute(this, name, value);
	}

	get reviewAuthor(): string {
		return this.core.reviewAuthor;
	}
	set reviewAuthor(value: string) {
		this.core.reviewAuthor = value || DEFAULT_REVIEW_AUTHOR;
		reflectAttribute(this, 'review-author', this.core.reviewAuthor);
	}

	/** File name shown in the title bar and used for downloads from the built-in File commands. */
	get fileName(): string {
		const { chrome } = this.core.shell;
		return chrome?.fileName ?? this.core.pendingFileName ?? DEFAULT_FILE_NAME;
	}
	set fileName(value: string) {
		const { chrome } = this.core.shell;
		if (chrome) chrome.fileName = value;
		else this.core.pendingFileName = value;
		reflectAttribute(this, 'file-name', this.fileName);
	}

	/** `light`, `dark`, or `auto` (default, inherits the Office theme, then follows the OS). Reflected to the `theme` attribute. */
	get theme(): EditorThemeMode {
		return this.core.theme;
	}
	set theme(value: EditorThemeMode) {
		const previous = this.core.theme;
		this.core.theme = normalizeThemeMode(value);
		reflectAttribute(this, 'theme', this.core.theme);
		if (this.core.theme !== 'auto') this.setAttribute('data-office-theme', this.core.theme);
		else if (previous !== 'auto' && this.getAttribute('data-office-theme') === previous)
			this.removeAttribute('data-office-theme');
	}
	/** Token overrides applied as inline `--dve-*` custom properties on the host. */
	get themeColors(): Partial<EditorTheme> | undefined {
		return this.core.themeColors;
	}
	set themeColors(value: Partial<EditorTheme> | undefined) {
		this.core.themeColors = value;
		this.core.appliedThemeVars = applyThemeColors(this, this.core.appliedThemeVars, value);
	}

	get locale(): string {
		return this.core.locale;
	}
	set locale(value: string) {
		const { core } = this;
		const { toolbar, searchPanel, review, chrome, canvas, paper } = core.shell;
		core.locale = normalizeEditorLocale(value);
		if (toolbar) setRibbonLocale(toolbar, core.locale);
		searchPanel?.setLocale(core.locale);
		core.parts.render(canvas, paper);
		review?.setLocale(core.locale);
		chrome?.setLocale(core.locale);
		core.inserts.setLocale(core.locale);
		core.formatDialogs.setLocale(core.locale);
		applyShellLabels(this, paper, core.locale);
		core.refreshControls();
		core.collab.presence?.relocalize();
		if (core.collab.yjs && core.view) core.view.dispatch(refreshWordYjsPresence(core.view.state));
		reflectAttribute(this, 'locale', core.locale);
	}

	get documentModel() {
		return this.core.model;
	}
	set documentModel(value: DocumentModel | null) {
		this.assertDocumentReplaceable();
		this.replaceDocument(value || createDocument());
		this.core.shell.chrome?.setSaveState('saved');
		if (this.isConnected) renderDocument(this.core);
		this.core.room.resume();
	}

	get readOnly() {
		return this.core.readOnly;
	}
	set readOnly(value: boolean) {
		this.core.readOnly = Boolean(value);
		this.core.applyReadOnly();
	}

	/**
	 * Shares the loaded document in a room while set; `null` leaves (see
	 * `DocxCollaborationOptions`). A document loaded later is shared too: the editor leaves
	 * before it is replaced and rejoins after. With `intent: 'join'` an empty room is never seeded.
	 * Race: with the default `create`, two windows that open one empty room at the same moment both
	 * seed it and the room holds both bodies, so give exactly one window `create` and the others
	 * `join`. A room name is also the Word `documentId`: nothing verifies that every window
	 * loaded the same source package. A viewer `user.role` blocks local edits (advisory).
	 * Leaving keeps the content; a session the editor opened is destroyed, a host `session` is not.
	 */
	get collaboration(): DocxCollaborationOptions | null {
		return this.core.room.wanted ?? null;
	}
	set collaboration(value: DocxCollaborationOptions | null | undefined) {
		this.core.room.want(value ?? undefined);
	}

	connectedCallback() {
		installOfficeUiThemeFor(this);
		buildShell(this.core, this.shellApi);
		renderDocument(this.core);
		this.core.room.resume();
	}

	disconnectedCallback() {
		const { core } = this;
		core.shell.surfaceNavigation?.();
		delete core.shell.surfaceNavigation;
		core.shell.printStoryEditing?.();
		core.shell.printNoteEditing?.();
		core.shell.printFloats?.destroy();
		delete core.shell.printFloats;
		core.shell.printEditing?.destroy();
		delete core.shell.printEditing;
		core.collab.presence?.leave();
		if (!core.collab.yjs) core.room.suspend();
		core.detachedState = core.view?.state;
		core.view?.destroy();
		core.view = undefined;
		core.imageMedia.release();
	}

	async load(input: Uint8Array | ArrayBuffer): Promise<void> {
		const { core } = this;
		this.assertDocumentReplaceable();
		const generation = ++core.loadGeneration;
		core.loading = true;
		core.refreshControls();
		try {
			const session = await loadDocument(input);
			if (generation !== core.loadGeneration) return;
			core.imageMedia.release();
			this.replaceDocument(session.model, session, false);
			core.shell.chrome?.setSaveState('saved');
			if (this.isConnected) renderDocument(core);
		} catch (cause) {
			const error = cause instanceof Error ? cause : new Error(String(cause));
			if (generation === core.loadGeneration) emit(this, 'document-error', error);
			throw error;
		} finally {
			if (generation === core.loadGeneration) {
				core.loading = false;
				core.refreshControls();
				core.room.resume();
			}
		}
	}

	setLoadedDocument(session: LoadedDocument) {
		this.assertDocumentReplaceable();
		this.replaceDocument(session.model, session);
		if (this.isConnected) renderDocument(this.core);
		this.core.room.resume();
	}

	/** Join only after loading the authority's matching document snapshot and version. */
	startCollaboration(config: CollaborationConfig): void {
		const { core } = this;
		if (core.collab.active) throw new Error('Stop the current collaboration session first.');
		if (!core.view) throw new Error('Mount and load the document before starting collaboration.');
		core.collab.start(config);
		core.loadGeneration++;
		core.detachedState = undefined;
		renderDocument(core);
	}

	getPendingCollaboration(): StepBatch | null {
		return this.core.pendingCollaboration();
	}

	receiveCollaboration(batch: unknown): ClientReceiveResult['status'] {
		const { collab, view } = this.core;
		if (!collab.client || !view) throw new Error('No mounted collaboration session.');
		const result = collab.client.receive(view.state, batch);
		if (result.status === 'applied') this.core.applyTransaction(result.transaction, true);
		return result.status;
	}

	/** Stopping with pending edits requires an explicit discard of the transport queue. */
	stopCollaboration(discardPending = false): void {
		const { core } = this;
		const state: EditorState | undefined = core.view?.state ?? core.detachedState;
		if (!discardPending && state && core.collab.client?.pendingStepCount(state))
			throw new Error(
				'Acknowledge pending collaboration edits before stopping, or explicitly discard the queue.',
			);
		core.room.stop();
		leaveYjs(core);
	}

	private assertDocumentReplaceable() {
		this.core.room.suspend();
		if (this.core.collab.active)
			throw new Error('Stop collaboration before replacing the document.');
	}

	/** Swaps in a new model (and the session that produced it) and invalidates in-flight loads. */
	private replaceDocument(model: DocumentModel, session?: LoadedDocument, bumpGeneration = true) {
		const { core } = this;
		core.detachedState = undefined;
		if (bumpGeneration) core.loadGeneration++;
		if (bumpGeneration) core.loading = false;
		core.loaded = session;
		core.model = model;
		core.inserts.reset();
		core.formatDialogs.closeAll();
		core.dirtyState.set(false);
	}
}

declare global {
	interface HTMLElementTagNameMap {
		'docx-editor': DocxEditorElement;
	}
}

export function registerDocxEditor() {
	registerSmartArt(); // no-op without a DOM and when already defined
	if (typeof customElements === 'undefined' || typeof HTMLElement === 'undefined') return;
	if (!customElements.get('docx-editor')) customElements.define('docx-editor', DocxEditorElement);
}

registerDocxEditor();
