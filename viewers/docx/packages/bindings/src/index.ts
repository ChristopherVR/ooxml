import {
	registerDocxEditor,
	type DocxCollaborationOptions,
	type DocxCollaborationState,
	type DocxEditorElement,
	type DocxEditorEventDetail,
	type DocxEditorEventName,
	type EditorLocaleInput,
	type EditorTheme,
	type EditorThemeMode,
	type FileCommandDetail,
	type PageChangeDetail,
	normalizeRibbonActions,
	type RibbonAction,
	type RibbonActionInput,
	type RibbonAddInCommandDetail,
	type RibbonAddInTab,
	type SurfaceHostActions,
} from 'docx-web-component';
import type { DocumentModel } from 'docx-core';
import { createPropTracker, sameThemeColors } from './forwarding';

/** Editor state a framework passes down as props. */
export interface EditorProps {
	// Framework props are `T | undefined` when unset, so `undefined` means "not provided".
	documentModel?: DocumentModel | undefined;
	readOnly?: boolean | undefined;
	/** Interface language: `en`, `fr`, `de`, `es`, `zh-CN`, or a tag such as `de-DE` that maps to one. */
	locale?: EditorLocaleInput | undefined;
	/** `light`, `dark`, or `auto` (default) to follow the OS color scheme. */
	theme?: EditorThemeMode | undefined;
	/**
	 * Name shown in the title bar and used by Save and Download (default `Document1.docx`). Forwarded
	 * when the prop changes, so File > Open can rename the document without being reverted.
	 */
	fileName?: string | undefined;
	/** Author recorded on tracked changes and comments. Default `Author`. Forwarded when it changes. */
	reviewAuthor?: string | undefined;
	/** Token overrides applied as `--dve-*` custom properties. Forwarded when its values change. */
	themeColors?: Partial<EditorTheme> | undefined;
	/** Left rail of page thumbnails (needs Print Layout). Default false. */
	showThumbnails?: boolean | undefined;
	/** Ribbon visibility. Default true. */
	showToolbar?: boolean | undefined;
	/** Ribbon controls to hide, by stable id (`RIBBON_ACTION_IDS`). Old English labels still work but warn. */
	hiddenActions?: readonly RibbonActionInput[] | undefined;
	/** Tabs the host adds after Word's own, as an Office add-in does. */
	ribbonAddIns?: readonly RibbonAddInTab[] | undefined;
	/**
	 * Shares the loaded document in a room while set; `null` leaves (see `DocxCollaborationOptions`).
	 * Forwarded only when the object changes, so pass a stable object, not a new literal per render.
	 */
	collaboration?: DocxCollaborationOptions | null | undefined;
}
/** Editor callbacks; each is the framework-neutral form of one entry in `EDITOR_EVENT_NAMES`. */
export interface EditorEventOptions {
	onDocumentChange?: ((model: DocumentModel) => void) | undefined;
	onDocumentError?: ((error: Error) => void) | undefined;
	onPageChange?: ((detail: PageChangeDetail) => void) | undefined;
	onDirtyChange?: ((dirty: boolean) => void) | undefined;
	/**
	 * A File command (new, open, save, saveAs, export, exportText, print, share) is about to run. The
	 * event is cancelable: call `event.preventDefault()` to take the command over yourself.
	 */
	onFileCommand?: ((event: CustomEvent<FileCommandDetail>) => void) | undefined;
	/** A non-fatal notice, such as a print layout that could not be produced faithfully. */
	onDocumentWarning?: ((message: string) => void) | undefined;
	/**
	 * The user toggled Editing / Viewing, `readOnly` changed or a collaboration role changed.
	 * `readOnly` is the effective value. A parent that syncs the `readOnly` prop from this callback
	 * must store `detail.hostReadOnly`, never `readOnly`, or a `viewer` role lock would be
	 * written into the host's setting.
	 */
	onReadOnlyChange?:
		| ((
				readOnly: boolean,
				detail: { readOnly: boolean; locked: boolean; hostReadOnly: boolean },
		  ) => void)
		| undefined;
	/** A ribbon control was activated. */
	onRibbonAction?: ((detail: RibbonAction) => void) | undefined;
	/** The user changed which ribbon commands are shown (File > Customize Ribbon); persist the ids. */
	onRibbonCustomize?: ((ids: readonly string[]) => void) | undefined;
	/** A command of a host tab (`ribbonAddIns`) was chosen. */
	onRibbonAddIn?: ((detail: RibbonAddInCommandDetail) => void) | undefined;
	/** Sharing started, stopped, reconnected or the people in the room changed. */
	onCollaborationChange?: ((state: DocxCollaborationState) => void) | undefined;
}
export interface EditorOptions extends EditorProps, EditorEventOptions {}

/** Single source of truth for the prop keys every adapter forwards. */
export const EDITOR_PROP_KEYS = [
	'documentModel',
	'readOnly',
	'fileName',
	'reviewAuthor',
	'themeColors',
	'locale',
	'theme',
	'showThumbnails',
	'showToolbar',
	'hiddenActions',
	'ribbonAddIns',
	'collaboration',
] as const satisfies readonly (keyof EditorProps)[];
export type EditorPropKey = (typeof EDITOR_PROP_KEYS)[number];
// Compile-time guard: adding a key to EditorProps without listing it above is an error.
const propKeysAreComplete: Exclude<keyof EditorProps, EditorPropKey> extends never ? true : never =
	true;
void propKeysAreComplete;

/** Single source of truth for the element events every adapter surfaces. */
export const EDITOR_EVENT_NAMES = [
	'document-change',
	'document-error',
	'document-warning',
	'page-change',
	'dirty-change',
	'readonly-change',
	'file-command',
	'ribbon-action',
	'ribbon-customize',
	'office-ribbon-add-in',
	'collaboration-change',
] as const satisfies readonly DocxEditorEventName[];
export type EditorEventName = (typeof EDITOR_EVENT_NAMES)[number];
/**
 * Handler for one bound event. `file-command` is cancelable, so its handler receives the event
 * itself (its `detail` is the command); every other handler receives the unwrapped detail.
 */
type EditorEventHandler<K extends EditorEventName> = K extends 'file-command'
	? (event: CustomEvent<DocxEditorEventDetail<K>>) => void
	: K extends 'readonly-change'
		? (readOnly: boolean, detail: DocxEditorEventDetail<K>) => void
		: (detail: DocxEditorEventDetail<K>) => void;
/** One handler per bound event; a missing key is a compile error in every adapter. */
export type EditorEventHandlers = {
	[K in EditorEventName]: EditorEventHandler<K> | undefined;
};
/** The DOM listener for each bound event, which re-emits it to the current options. */
type EditorListeners = {
	[K in EditorEventName]: (event: CustomEvent<DocxEditorEventDetail<K>>) => void;
};

function sameList(a: readonly string[], b: readonly string[]): boolean {
	return a.length === b.length && a.every((item, index) => item === b[index]);
}
function copyProp<K extends EditorPropKey>(to: EditorProps, from: EditorProps, key: K): void {
	to[key] = from[key];
}
/** Copies exactly the shared props out of a framework's props/instance object. */
export function pickEditorProps(source: EditorProps): EditorProps {
	const picked: EditorProps = {};
	for (const key of EDITOR_PROP_KEYS) copyProp(picked, source, key);
	return picked;
}
/** Maps per-event handlers (keyed by DOM event name) onto the shared option callbacks. */
export function eventOptions(handlers: EditorEventHandlers): EditorEventOptions {
	return {
		onDocumentChange: handlers['document-change'],
		onDocumentError: handlers['document-error'],
		onPageChange: handlers['page-change'],
		onDirtyChange: handlers['dirty-change'],
		onFileCommand: handlers['file-command'],
		onDocumentWarning: handlers['document-warning'],
		onReadOnlyChange: handlers['readonly-change'],
		onRibbonAction: handlers['ribbon-action'],
		onRibbonCustomize: handlers['ribbon-customize'],
		onRibbonAddIn: handlers['office-ribbon-add-in'],
		onCollaborationChange: handlers['collaboration-change'],
	};
}

export interface EditorHandle {
	readonly element: DocxEditorElement;
	load(input: Uint8Array | ArrayBuffer): Promise<void>;
	/** The saved document as a Blob. Does not clear `dirty`; call `markClean()` after persisting it. */
	save(): Promise<Blob>;
	/** Saves and downloads in the browser, then marks the document clean. */
	download(fileName?: string): Promise<void>;
	markClean(): void;
	readonly dirty: boolean;
}
export interface EditorBinding extends EditorHandle {
	update(options: EditorOptions): void;
	destroy(): void;
}
/** All framework adapters share property, event and lifecycle semantics here. */
export function mountEditor(
	host: HTMLElement,
	initial: EditorOptions = {},
	surfaceOnly = false,
): EditorBinding {
	registerDocxEditor();
	const element = host.ownerDocument.createElement('docx-editor');
	if (surfaceOnly && typeof element.getToolbarProps !== 'function')
		throw new Error('Update ooxml-ui to use editor building blocks.');
	element.surfaceOnly = surfaceOnly;
	let options: EditorOptions = {};
	let lastInput: DocumentModel | undefined;
	let lastEmitted: DocumentModel | undefined;
	let lastAddIns: readonly RibbonAddInTab[] | undefined;
	let lastCollaboration: DocxCollaborationOptions | null = null;
	let destroyed = false;
	const changedProp = createPropTracker<EditorPropKey>();
	const listeners: EditorListeners = {
		'document-change': (event) => {
			lastEmitted = event.detail;
			options.onDocumentChange?.(lastEmitted);
		},
		'document-error': (event) => options.onDocumentError?.(event.detail),
		'document-warning': (event) => options.onDocumentWarning?.(event.detail),
		'page-change': (event) => options.onPageChange?.(event.detail),
		'dirty-change': (event) => options.onDirtyChange?.(event.detail),
		'readonly-change': (event) => options.onReadOnlyChange?.(event.detail.readOnly, event.detail),
		'file-command': (event) => options.onFileCommand?.(event),
		'ribbon-action': (event) => options.onRibbonAction?.(event.detail),
		'ribbon-customize': (event) => options.onRibbonCustomize?.(event.detail),
		'office-ribbon-add-in': (event) => options.onRibbonAddIn?.(event.detail),
		'collaboration-change': (event) => options.onCollaborationChange?.(event.detail),
	};
	const listen = (add: boolean, names: readonly EditorEventName[] = EDITOR_EVENT_NAMES) => {
		for (const name of names) {
			const listener = listeners[name] as EventListener;
			if (add) element.addEventListener(name, listener);
			else element.removeEventListener(name, listener);
		}
	};
	listen(true);
	const binding: EditorBinding = {
		element,
		update(next) {
			if (destroyed) return;
			options = next;
			element.locale = next.locale ?? 'en';
			// Forwarded only when the prop changes: a parent that does not sync `readOnly` from
			// `onReadOnlyChange` must not revert the user's Editing / Viewing toggle on re-render.
			if (changedProp('readOnly', next.readOnly ?? false))
				element.readOnly = next.readOnly ?? false;
			// Optional values are forwarded only when the parent sets or changes them, so a re-render
			// never undoes a rename or review author the user chose inside the editor.
			if (next.fileName !== undefined && changedProp('fileName', next.fileName))
				element.fileName = next.fileName;
			if (next.reviewAuthor !== undefined && changedProp('reviewAuthor', next.reviewAuthor))
				element.reviewAuthor = next.reviewAuthor;
			if (
				next.themeColors !== undefined &&
				changedProp('themeColors', next.themeColors, sameThemeColors)
			)
				element.themeColors = { ...next.themeColors };
			element.theme = next.theme ?? 'auto';
			element.showThumbnails = next.showThumbnails ?? false;
			element.showToolbar = next.showToolbar ?? true;
			if (!sameList(element.hiddenActions, normalizeRibbonActions(next.hiddenActions ?? []).ids))
				element.hiddenActions = next.hiddenActions ?? [];
			// The element keeps panels whose look is unchanged, so a fresh array per render is cheap.
			if (next.ribbonAddIns !== lastAddIns) element.ribbonAddIns = next.ribbonAddIns ?? [];
			lastAddIns = next.ribbonAddIns;
			// Before the document model: a document opened below joins the room.
			if ((next.collaboration ?? null) !== lastCollaboration) {
				lastCollaboration = next.collaboration ?? null;
				element.collaboration = lastCollaboration;
			}
			if (
				next.documentModel &&
				next.documentModel !== lastInput &&
				next.documentModel !== lastEmitted
			) {
				element.documentModel = next.documentModel;
			}
			lastInput = next.documentModel;
		},
		load: (input) => element.load(input),
		save: () => element.save(),
		download: (fileName) => element.download(fileName),
		markClean: () => element.markClean(),
		get dirty() {
			return element.dirty;
		},
		destroy() {
			if (destroyed) return;
			destroyed = true;
			// The collaboration listener stays until the room is left, so its "stopped" change reaches the host.
			listen(
				false,
				EDITOR_EVENT_NAMES.filter((name) => name !== 'collaboration-change'),
			);
			// Leave a room this binding joined: an unmounted editor must not keep editing the shared
			// document. A session the host started itself (startYjsCollaboration) stays the host's.
			if (lastCollaboration) element.stopCollaboration();
			listen(false, ['collaboration-change']);
			element.remove();
		},
	};
	binding.update(initial);
	host.append(element);
	return binding;
}

/**
 * The imperative handle every framework exposes (React ref, Vue template ref, Angular component
 * instance, Solid `editorRef`, Svelte component instance). `handle-keys.test.ts` checks each one.
 */
export const EDITOR_HANDLE_KEYS = [
	'element',
	'load',
	'save',
	'download',
	'markClean',
	'dirty',
] as const satisfies readonly (keyof EditorHandle)[];
export type EditorHandleKey = (typeof EDITOR_HANDLE_KEYS)[number];
// Compile-time guard: adding a member to EditorHandle without listing it above is an error.
const handleKeysAreComplete: Exclude<keyof EditorHandle, EditorHandleKey> extends never
	? true
	: never = true;
void handleKeysAreComplete;

/** Ribbon and chrome options that a surface, which draws none, would silently ignore. */
type RibbonOnlyOption =
	| 'showToolbar'
	| 'showThumbnails'
	| 'hiddenActions'
	| 'ribbonAddIns'
	| 'onRibbonAddIn';
/**
 * Framework-neutral composition API, shared by every adapter. Choose before mount. Ribbon-only
 * options are not part of it. The `onOpen...` callbacks are the host's own dialogs: a workflow
 * without its callback reports as not enabled.
 */
export interface SurfaceOptions extends Omit<EditorOptions, RibbonOnlyOption>, SurfaceHostActions {
	canEdit?: boolean | undefined;
}
export function mountEditorSurface(host: HTMLElement, initial: SurfaceOptions = {}) {
	const effective = (options: SurfaceOptions): EditorOptions => ({
		...options,
		readOnly: options.canEdit !== true || options.readOnly === true,
	});
	const actions = (options: SurfaceOptions): SurfaceHostActions => ({
		...(options.onOpenHyperlinkDialog && { onOpenHyperlinkDialog: options.onOpenHyperlinkDialog }),
		...(options.onOpenFindReplace && { onOpenFindReplace: options.onOpenFindReplace }),
		...(options.onOpenImagePicker && { onOpenImagePicker: options.onOpenImagePicker }),
	});
	const binding = mountEditor(host, effective(initial), true);
	const update = binding.update;
	const applyActions = (options: SurfaceOptions) => {
		const next = actions(options);
		const current = binding.element.surfaceActions;
		const keys = new Set([
			...Object.keys(current),
			...Object.keys(next),
		] as (keyof SurfaceHostActions)[]);
		if ([...keys].some((key) => current[key] !== next[key])) binding.element.surfaceActions = next;
	};
	applyActions(initial);
	return Object.assign(binding, {
		update: (options: SurfaceOptions) => {
			update(effective(options));
			applyActions(options);
		},
		getToolbarProps: () => binding.element.getToolbarProps(),
		getNavigationProps: () => binding.element.getNavigationProps(),
		getStatusProps: () => binding.element.getStatusProps(),
		subscribe: (listener: () => void) => binding.element.subscribeEditorState(listener),
	});
}
