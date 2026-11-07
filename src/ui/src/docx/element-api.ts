import type { CloudFileProvider, CloudOpenedFile, DocumentStorage } from 'ooxml-core/embedding';
import type { RibbonAddInTab } from '../ribbon/add-in-tabs';
import { applyRibbonAddIns } from './editor-shell';
import type { CollabSession } from 'ooxml-core/collab';
import type { WordYjsOptions } from 'ooxml-core/docx/ui';
import { joinYjs } from './collaboration-room';
import type { DocxCollaborationState } from './collaboration-types';
import { yjsPresenceProfile } from './yjs-presence';
import { saveDocx } from 'ooxml-core/docx';
import { schema } from './schema';
import { docToModel } from './model-adapter';
import type { EditorCore } from './editor-core';
import { modifiedStamp, reflectAttribute } from './editor-attributes';
import { docxFileName, downloadBytes, wordBlob } from './file-commands';
import {
	normalizeRibbonActions,
	type RibbonActionId,
	type RibbonActionInput,
} from 'ooxml-core/docx/ui';
import { emit } from './events';
import { applyViewOptions } from './view-options';
import type { WordCustomization } from './customization';
import { setWordCustomization } from './customization-apply';
import { wordToolbarProps, type WordToolbarProps } from './building-blocks';
import {
	wordNavigationProps,
	wordStatusProps,
	type WordNavigationProps,
	type WordStatusProps,
} from './building-blocks-status';
import type { SurfaceHostActions } from './building-blocks-actions';
import type { GoogleFontsOption } from 'ooxml-core/text';
import { refreshDocxWebfonts } from './docx-webfonts';

const HTMLElementBase: typeof HTMLElement =
	typeof HTMLElement === 'undefined' ? (class {} as typeof HTMLElement) : HTMLElement;

/**
 * Save, collaboration, dirty-tracking and UI-customisation surface of `<docx-editor>`.
 * the shared `EditorCore`, so the element class stays a thin lifecycle facade.
 */
export abstract class DocxEditorApi extends HTMLElementBase {
	/** Choose before connecting. This mode mounts only the document and its editing surface. */
	get surfaceOnly(): boolean {
		return this.core.surfaceOnly;
	}
	set surfaceOnly(value: boolean) {
		if (this.core.shell.paper && value !== this.core.surfaceOnly)
			throw new Error('Choose surfaceOnly before mounting the editor.');
		this.core.surfaceOnly = Boolean(value);
	}
	getToolbarProps(): WordToolbarProps {
		return wordToolbarProps(this.core);
	}
	/** Page, view mode and navigation callbacks for custom page controls. */
	getNavigationProps(): WordNavigationProps {
		return wordNavigationProps(this.core);
	}
	/** Word count and zoom for custom status controls. */
	getStatusProps(): WordStatusProps {
		return wordStatusProps(this.core);
	}
	/**
	 * Host-owned dialogs for surface mode: `onOpenHyperlinkDialog`, `onOpenFindReplace` and
	 * `onOpenImagePicker`. A workflow without its callback is reported as not enabled.
	 */
	get surfaceActions(): SurfaceHostActions {
		return this.core.surfaceActions;
	}
	set surfaceActions(value: SurfaceHostActions) {
		this.core.surfaceActions = { ...value };
		if (this.core.view) this.core.refreshControls();
	}
	/**
	 * Host-owned storage. When set, File > Save and Save As hand the saved bytes to
	 * `storage.save()` instead of downloading them, and File > Open asks `storage.open()` (when it
	 * has one) instead of showing the file picker. The cancelable `file-command` event still fires
	 * first; a host that prevents it handles the command itself.
	 */
	get storage(): DocumentStorage | undefined {
		return this.core.storage;
	}
	set storage(value: DocumentStorage | null | undefined) {
		this.core.storage = value ?? undefined;
	}
	/**
	 * Cloud places for File > Open (OneDrive and the like). When non-empty, Open lists "This
	 * device" and one place per provider. A file opened from a place loads like any other, locks
	 * the editor when the provider says it cannot be edited, and saves through the file's own
	 * storage (not `storage`) until another document loads. Opening from a place does not fire
	 * `file-command`: the host chose the providers.
	 */
	get fileProviders(): readonly CloudFileProvider[] {
		return this.core.fileProviders;
	}
	set fileProviders(value: readonly CloudFileProvider[] | null | undefined) {
		const next = value ?? [];
		const current = this.core.fileProviders;
		if (next.length === current.length && next.every((p, i) => p === current[i])) return;
		this.core.fileProviders = [...next];
		this.core.shell.chrome?.backstage.relocalize();
	}
	/**
	 * Webfonts for the document fonts the reader has neither installed nor embedded. Unset (the
	 * default): the reader decides in File > Options > Trust Center (off until turned on; when on,
	 * Google Fonts). `false`: off, nothing is requested. A set value locks that toggle, which then
	 * shows it read-only. `{ baseUrl }`: such a family that the Google
	 * Fonts catalogue serves, or its metric-compatible clone (Calibri as Carlito), is requested from
	 * that CSS2 endpoint (`https://fonts.googleapis.com/css2`, or a self-hosted mirror) with only
	 * the family names, and `load` waits briefly (150 ms) for the faces before showing the document. An
	 * invalid `baseUrl` warns and stays off. Read when a document loads; changing it re-requests
	 * for the document on screen.
	 */
	get googleFonts(): GoogleFontsOption | undefined {
		return this.core.webfonts.option;
	}
	set googleFonts(value: GoogleFontsOption | undefined) {
		this.core.webfonts.option = value;
		if (this.core.view) void refreshDocxWebfonts(this.core);
	}
	get loading(): boolean {
		return this.core.loading;
	}
	/** Selection and formatting updates, including changes which do not dirty the document. */
	subscribeEditorState(listener: () => void): () => void {
		this.addEventListener('editor-state-change', listener);
		return () => this.removeEventListener('editor-state-change', listener);
	}
	protected abstract readonly core: EditorCore;
	abstract fileName: string;
	abstract reviewAuthor: string;

	/** Join a synchronized Yjs room after loading its matching source package. */
	startYjsCollaboration(session: CollabSession, options: WordYjsOptions): void {
		joinYjs(this.core, session, options);
	}

	/** Who is in the room, and whether it is synced; the off state when not sharing. */
	get collaborationState(): DocxCollaborationState {
		return this.core.room.state();
	}

	reconnectCollaboration(): void {
		const { core } = this;
		const session = core.collab.yjs?.session ?? core.room.pendingSession;
		if (!session) throw new Error('No Yjs collaboration session.');
		session.reconnect();
	}

	resyncCollaboration(): boolean {
		return this.core.collab.yjs?.resync() ?? false;
	}

	publishPresence(profile: { name: string; color: string }) {
		if (this.core.collab.yjs) {
			const validated = yjsPresenceProfile(profile);
			this.reviewAuthor = validated.name;
			this.core.collab.yjs.session.awareness.setLocalStateField('user', validated);
			return null;
		}
		if (!this.core.collab.presence)
			throw new Error('Start collaboration before publishing presence.');
		this.reviewAuthor = profile.name || this.reviewAuthor;
		return this.core.collab.presence.publish(profile);
	}
	receivePresence(message: unknown) {
		if (!this.core.collab.presence)
			throw new Error('Start collaboration before receiving presence.');
		return this.core.collab.presence.receive(message);
	}
	leavePresence() {
		this.core.collab.yjs?.session.awareness.setLocalStateField('cursor', null);
		return this.core.collab.presence?.leave() ?? null;
	}

	/** Shows the left rail of page thumbnails. Needs Print Layout; reflected to `show-thumbnails`. */
	get showThumbnails(): boolean {
		return this.core.viewOptions.showThumbnails;
	}
	set showThumbnails(value: boolean) {
		const next = Boolean(value);
		if (next !== this.core.viewOptions.showThumbnails) {
			this.core.viewOptions.showThumbnails = next;
			applyViewOptions(this.core);
		}
		reflectAttribute(this, 'show-thumbnails', this.showThumbnails);
	}

	/**
	 * Tabs a host adds after Word's own, as an Office add-in does. Each command runs its `run`
	 * callback and dispatches `office-ribbon-add-in` (`{ tab, command }`) from the editor.
	 */
	get ribbonAddIns(): readonly RibbonAddInTab[] {
		return this.core.ribbonAddIns;
	}
	set ribbonAddIns(value: readonly RibbonAddInTab[]) {
		this.core.ribbonAddIns = value ?? [];
		applyRibbonAddIns(this.core);
	}

	/** Hides the ribbon (the title bar and status bar stay). Default true; `show-toolbar="false"`. */
	get showToolbar(): boolean {
		return this.core.viewOptions.showToolbar;
	}
	set showToolbar(value: boolean) {
		const next = Boolean(value);
		if (next !== this.core.viewOptions.showToolbar) {
			this.core.viewOptions.showToolbar = next;
			applyViewOptions(this.core);
		}
		reflectAttribute(this, 'show-toolbar', this.showToolbar);
	}

	/**
	 * Ribbon controls to hide, by stable id (`RIBBON_ACTION_IDS`, e.g. `'bold'`). Unknown ids are
	 * ignored. The old English labels (`'Bold'`) are still accepted for one release: they are mapped
	 * to ids and reported through a `document-warning` event. Reading returns ids only.
	 */
	get hiddenActions(): readonly RibbonActionId[] {
		return this.core.viewOptions.hiddenActions;
	}
	set hiddenActions(value: readonly RibbonActionInput[]) {
		const { ids: next, legacy } = normalizeRibbonActions(Array.isArray(value) ? value : []);
		if (legacy.length > 0)
			emit(
				this,
				'document-warning',
				`hiddenActions: English labels (${legacy.map((label) => `'${label}'`).join(', ')}) are deprecated and will be removed in the next release; use the stable ids from RIBBON_ACTION_IDS.`,
			);
		const current = this.core.viewOptions.hiddenActions;
		if (next.length === current.length && next.every((id, index) => id === current[index])) return;
		this.core.viewOptions.hiddenActions = next;
		applyViewOptions(this.core);
	}

	/**
	 * The host customisation (`WordCustomization`): `capabilities` (removing every UI path to
	 * what is switched off), `backstage.extraPages` (host File tab pages) and
	 * `contextMenu.extraCommands` (host context-menu commands). The object is applied when its
	 * identity changes, so keep it stable; `null` restores the defaults.
	 */
	get customization(): WordCustomization | null {
		return this.core.viewOptions.customizationInput;
	}
	set customization(value: WordCustomization | null | undefined) {
		setWordCustomization(this.core, value ?? null);
	}

	/** True after an edit until the document is saved through File > Save, `download()`, loaded or `markClean()`. */
	get dirty(): boolean {
		return this.core.dirtyState.dirty;
	}

	/**
	 * Opens a file a host got from a `CloudFileProvider` (its own picker or a start page), exactly as
	 * File > Open > OneDrive does: the file's name, read-only when `file.canEdit` is false, and Save
	 * routed to `file.storage` until another document loads. A load failure is reported through
	 * `document-error`. Does nothing when the host removed the Open command.
	 */
	async openCloudFile(file: CloudOpenedFile): Promise<void> {
		await this.core.shell.chrome?.openCloudFile(file);
	}

	/** Declares the current content saved, for hosts that persist the result of `save()` themselves. */
	markClean(): void {
		this.core.dirtyState.set(false);
		this.core.shell.chrome?.setSaveState('saved');
	}

	/** Serialized document bytes: the loaded package's own writer, else a fresh DOCX from the model. */
	async saveBytes(): Promise<Uint8Array> {
		const { core } = this;
		if (core.collab.yjs) {
			core.model = docToModel(core.collab.yjs.state(schema).doc, core.model);
			if (core.collab.yjs.sharedComments)
				core.model = { ...core.model, comments: core.collab.yjs.comments.all() };
		}
		// Only pass staged pictures when there are any: legacy DOC sessions take the model alone.
		const pending = new Map([
			...core.inserts.pendingMedia,
			...(core.collab.yjs?.media.all() ?? []),
		]);
		const media = pending.size ? pending : undefined;
		if (core.loaded) {
			// A converted Word 97-2003 .doc saves as .docx here; its `save()` keeps the binary format.
			const options = { modified: modifiedStamp(core.reviewAuthor) };
			return core.loaded.saveDocx
				? core.loaded.saveDocx(core.model, media, options)
				: core.loaded.save(core.model, media, options);
		}
		return saveDocx(core.model, media);
	}

	/** The saved document as a Blob. Does not change `dirty`; call `markClean()` after persisting it. */
	async save(): Promise<Blob> {
		return wordBlob(await this.saveBytes(), this.fileName);
	}

	/** Saves and starts a browser download, then marks the document clean. */
	async download(fileName?: string): Promise<void> {
		const name = fileName || this.fileName;
		downloadBytes(
			await this.saveBytes(),
			/\.[^./\\]+$/.test(name) ? docxFileName(name) : `${name}.docx`,
		);
		this.markClean();
	}
}
