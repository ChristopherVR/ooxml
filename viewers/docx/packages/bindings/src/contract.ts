// The Word binding contract: the props every adapter forwards to `<docx-editor>` and how, the
// element events it re-emits and the handle it exposes. The adapters are generated from or checked
// against `EDITOR_CONTRACT` (`contract-parity.test.ts`); the machinery is `ooxml-ui/bindings`.
import {
	normalizeRibbonActions,
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
	type RibbonAction,
	type RibbonActionInput,
	type RibbonAddInCommandDetail,
	type RibbonAddInTab,
	type WordCustomization,
} from 'docx-web-component';
import type { DocumentModel } from 'docx-core';
import type { DocumentStorage } from 'ooxml-core/embedding';
import {
	contractOptions,
	deferredHandle as kitDeferredHandle,
	eventArgs,
	eventOptions as mapEventOptions,
	pickProps,
	readOnlyChange,
	sameList,
	shallowEqual,
	svelteHandlerName,
	type BindingContract,
} from 'ooxml-ui/bindings';

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
	/**
	 * Host-owned storage: File > Save and Save As go to `storage.save()` instead of a download, and
	 * File > Open to `storage.open()` when it has one. `file-command` still fires first.
	 */
	storage?: DocumentStorage | null | undefined;
	/**
	 * Host customisation (`WordCustomization`): `capabilities`, `backstage.extraPages` and
	 * `contextMenu.extraCommands`. Forwarded when the object changes, so keep it stable (memoise
	 * it); `null` or omitted restores the defaults.
	 */
	customization?: WordCustomization | null | undefined;
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
	'storage',
	'customization',
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
 * Props in forwarding order. `readOnly` and the objects are forwarded only when they change, so a
 * parent that does not sync them never reverts what the user chose inside the editor; customisation
 * precedes the read-only state it can force, and collaboration precedes the document model
 * (forwarded by `mountEditor`), so a document opened in the same update joins the room.
 */
export const EDITOR_CONTRACT: BindingContract<
	DocxEditorElement,
	EditorProps,
	EditorHandle,
	EditorEventName,
	keyof EditorEventOptions
> = {
	props: {
		locale: { forward: 'always', fallback: 'en' },
		readOnly: { forward: 'changed', fallback: false },
		fileName: { forward: 'defined' },
		reviewAuthor: { forward: 'defined' },
		themeColors: {
			forward: 'defined',
			same: shallowEqual,
			apply: (element, value) => (element.themeColors = { ...value }),
		},
		theme: { forward: 'always', fallback: 'auto' },
		showThumbnails: { forward: 'always', fallback: false },
		showToolbar: { forward: 'always', fallback: true },
		hiddenActions: {
			forward: 'always',
			apply: (element, value) => {
				if (!sameList(element.hiddenActions, normalizeRibbonActions(value ?? []).ids))
					element.hiddenActions = value ?? [];
			},
		},
		// The element keeps panels whose look is unchanged, so a fresh array per render is cheap.
		ribbonAddIns: {
			forward: 'changed',
			initial: undefined,
			apply: (element, value) => (element.ribbonAddIns = value ?? []),
		},
		customization: { forward: 'changed', fallback: null },
		storage: {
			forward: 'changed',
			fallback: null,
			apply: (element, value) => (element.storage = value ?? undefined),
		},
		collaboration: { forward: 'changed', fallback: null, initial: null },
		documentModel: { forward: 'manual' },
	},
	events: {
		'document-change': { option: 'onDocumentChange' },
		'document-error': { option: 'onDocumentError' },
		'document-warning': { option: 'onDocumentWarning' },
		'page-change': { option: 'onPageChange' },
		'dirty-change': { option: 'onDirtyChange' },
		'readonly-change': readOnlyChange,
		// Cancelable: the handler gets the event itself, so a host can call preventDefault().
		'file-command': { option: 'onFileCommand', args: eventArgs.event },
		'ribbon-action': { option: 'onRibbonAction' },
		'ribbon-customize': { option: 'onRibbonCustomize' },
		'office-ribbon-add-in': { option: 'onRibbonAddIn' },
		'collaboration-change': { option: 'onCollaborationChange' },
	},
	handle: {
		element: 'get',
		load: 'async',
		save: 'async',
		download: 'async',
		markClean: 'optional-call',
		dirty: { fallback: false },
	},
};

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

/** Copies exactly the shared props out of a framework's props/instance object. */
export function pickEditorProps(source: EditorProps): EditorProps {
	return pickProps<EditorProps>(EDITOR_CONTRACT, source);
}
/** Maps per-event handlers (keyed by DOM event name) onto the shared option callbacks. */
export function eventOptions(handlers: EditorEventHandlers): EditorEventOptions {
	return mapEventOptions<EditorEventOptions>(EDITOR_CONTRACT, handlers);
}

/** Options from Svelte's props, whose callbacks are the lower-case options (`ondocumentchange`). */
export function svelteEditorOptions(props: object): EditorOptions {
	return contractOptions<EditorOptions>(EDITOR_CONTRACT, props, svelteHandlerName);
}
/**
 * A handle for adapters that hold a binding which may not be mounted yet: every call reads the
 * current binding and throws a clear error before mount.
 */
export function deferredHandle(current: () => EditorHandle | null | undefined): EditorHandle {
	return kitDeferredHandle(EDITOR_CONTRACT.handle, current);
}

/**
 * The imperative handle every framework exposes (React ref, Vue template ref, Angular component
 * instance, Solid `editorRef`, Svelte component instance). `contract-parity.test.ts` checks each.
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
