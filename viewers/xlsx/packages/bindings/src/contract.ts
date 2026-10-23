// The Excel binding contract: the props every adapter forwards to `<xlsx-editor>` and how, the
// element events it re-emits and the handle it exposes. The adapters are generated from or checked
// against `EDITOR_CONTRACT` (`contract-parity.test.ts`); the machinery is `ooxml-ui/bindings`.
import type {
	EditorLocaleInput,
	EditorThemeMode,
	FileCommandDetail,
	RibbonAddInCommandDetail,
	RibbonAddInTab,
	SelectionChangeDetail,
	XlsxCollaborationOptions,
	XlsxCollaborationState,
	XlsxEditorElement,
	XlsxEditorEventName,
	XlsxThemeColors,
	ExcelCustomization,
} from 'xlsx-web-component';
import type { Workbook } from '@christophervr/xlsx-core';
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
	/** A workbook model to show. Re-assigned only when a new object is passed (never echoed back). */
	workbook?: Workbook | undefined;
	/** File bytes to open (.xlsx, .xlsm, .xls, .csv); loaded again whenever a new array is passed. */
	bytes?: Uint8Array | ArrayBuffer | undefined;
	/** URL of a file to fetch and open; fetched again when the string changes. */
	src?: string | undefined;
	/** Name shown in the title bar and used by Save and Download. Default `Book1.xlsx`. */
	fileName?: string | undefined;
	readOnly?: boolean | undefined;
	/** Interface language: `en`, `fr`, `de`, `es`, `zh-CN`, or a tag such as `de-DE` that maps to one. */
	locale?: EditorLocaleInput | undefined;
	/** `light`, `dark`, or `auto` (default) to follow the OS color scheme. */
	theme?: EditorThemeMode | undefined;
	/** Author recorded on new comments. Default `Author`. */
	authorName?: string | undefined;
	/** Ribbon visibility. Default true. */
	showToolbar?: boolean | undefined;
	/** Formula bar and name box visibility. Default true. */
	showFormulaBar?: boolean | undefined;
	/** Ribbon controls to hide, by stable id. */
	hiddenActions?: readonly string[] | undefined;
	/** Tabs the host adds after Excel's own, as an Office add-in does. */
	ribbonAddIns?: readonly RibbonAddInTab[] | undefined;
	/** Theme token overrides, applied as `--xve-*` custom properties. */
	themeColors?: XlsxThemeColors | undefined;
	/**
	 * Shares the open workbook in a room while set (room id, server URL or provider, user name and
	 * colour); forwarded when a new object is passed, and `undefined` or `null` leaves the room.
	 */
	collaboration?: XlsxCollaborationOptions | null | undefined;
	/**
	 * Host-owned storage: File > Save and Save As go to `storage.save()` instead of a download, and
	 * File > Open to `storage.open()` when it has one. `file-command` still fires first.
	 */
	storage?: DocumentStorage | null | undefined;
	/**
	 * Host customisation (`ExcelCustomization`): `capabilities`, `backstage.extraPages` and
	 * `contextMenu.extraCommands`. Forwarded when the object changes, so keep it stable (memoise
	 * it); `null` or omitted restores the defaults.
	 */
	customization?: ExcelCustomization | null | undefined;
}
/** Editor callbacks; each is the framework-neutral form of one entry in `EDITOR_EVENT_NAMES`. */
export interface EditorEventOptions {
	onWorkbookChange?: ((workbook: Workbook) => void) | undefined;
	onWorkbookError?: ((error: Error) => void) | undefined;
	onSelectionChange?: ((detail: SelectionChangeDetail) => void) | undefined;
	onDirtyChange?: ((dirty: boolean) => void) | undefined;
	/**
	 * The user toggled Editing / Viewing, `readOnly` changed or a collaboration role changed.
	 * `readOnly` is the effective value. A controlled parent that mirrors the switch must store
	 * `detail.hostReadOnly`, never `readOnly`, or a `viewer` role lock would be written into the
	 * host's setting.
	 */
	onReadOnlyChange?:
		| ((
				readOnly: boolean,
				detail: { readOnly: boolean; locked: boolean; hostReadOnly: boolean },
		  ) => void)
		| undefined;
	/**
	 * A File command (new, open, save, saveAs, export, exportCsv, print) is about to run. The event is
	 * cancelable: call `event.preventDefault()` to take the command over yourself.
	 */
	onFileCommand?: ((event: CustomEvent<FileCommandDetail>) => void) | undefined;
	/** Non-fatal notes from the loader (unsupported features, compatibility). */
	onWorkbookWarning?: ((warnings: string[]) => void) | undefined;
	/** A command ran from the ribbon, a menu, Tell me or a shortcut. */
	onRibbonAction?: ((detail: { id: string }) => void) | undefined;
	/** The user changed the shown ribbon commands (File > Options > Customize Ribbon). */
	onRibbonCustomize?: ((hiddenActions: string[]) => void) | undefined;
	/** The active sheet changed. */
	onSheetChange?: ((detail: { index: number; name: string }) => void) | undefined;
	/** A command of a host tab (`ribbonAddIns`) was chosen. */
	onRibbonAddIn?: ((detail: RibbonAddInCommandDetail) => void) | undefined;
	/** Sharing started or stopped, the connection changed, or someone joined or left. */
	onCollaborationChange?: ((state: XlsxCollaborationState) => void) | undefined;
	/** Called once, after the element is created and attached. */
	onReady?: ((element: XlsxEditorElement) => void) | undefined;
}
export interface EditorOptions extends EditorProps, EditorEventOptions {}

/** Single source of truth for the prop keys every adapter forwards. */
export const EDITOR_PROP_KEYS = [
	'workbook',
	'bytes',
	'src',
	'fileName',
	'readOnly',
	'locale',
	'theme',
	'authorName',
	'showToolbar',
	'showFormulaBar',
	'hiddenActions',
	'ribbonAddIns',
	'themeColors',
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
	'workbook-change',
	'workbook-error',
	'selection-change',
	'dirty-change',
	'readonly-change',
	'file-command',
	'workbook-warning',
	'ribbon-action',
	'ribbon-customize',
	'sheet-change',
	'office-ribbon-add-in',
	'collaboration-change',
] as const satisfies readonly XlsxEditorEventName[];
export type EditorEventName = (typeof EDITOR_EVENT_NAMES)[number];
/** One handler per bound event (unwrapped payload); a missing key is a compile error in every adapter. */
export interface EditorEventHandlers {
	'workbook-change': EditorEventOptions['onWorkbookChange'];
	'workbook-error': EditorEventOptions['onWorkbookError'];
	'selection-change': EditorEventOptions['onSelectionChange'];
	'dirty-change': EditorEventOptions['onDirtyChange'];
	'readonly-change': EditorEventOptions['onReadOnlyChange'];
	'file-command': EditorEventOptions['onFileCommand'];
	'workbook-warning': EditorEventOptions['onWorkbookWarning'];
	'ribbon-action': EditorEventOptions['onRibbonAction'];
	'ribbon-customize': EditorEventOptions['onRibbonCustomize'];
	'sheet-change': EditorEventOptions['onSheetChange'];
	'office-ribbon-add-in': EditorEventOptions['onRibbonAddIn'];
	'collaboration-change': EditorEventOptions['onCollaborationChange'];
	ready?: EditorEventOptions['onReady'];
}

export interface EditorHandle {
	readonly element: XlsxEditorElement;
	load(input: Uint8Array | ArrayBuffer, fileName?: string): Promise<void>;
	newWorkbook(): void;
	/** The saved workbook as an .xlsx Blob. Does not clear `dirty`; call `markClean()` after persisting it. */
	save(): Promise<Blob>;
	saveBytes(format?: 'xlsx' | 'csv'): Promise<Uint8Array>;
	/** Saves and downloads in the browser, then marks the workbook clean. */
	download(fileName?: string): Promise<void>;
	markClean(): void;
	select(ref: string): void;
	getSelection(): string;
	setActiveSheet(index: number): void;
	/** Opens File > Share, where the user starts or stops sharing the workbook. */
	share(): void;
	readonly dirty: boolean;
}
export interface EditorBinding extends EditorHandle {
	update(options: EditorOptions): void;
	destroy(): void;
}

/**
 * Props in forwarding order; each is written only when the prop itself changed since it was last
 * written, so a re-render never undoes what the user changed inside the editor. `fileName` and
 * collaboration precede the workbook, bytes and `src` (forwarded by `mountEditor`), so a workbook
 * opened in the same update is named and joins the room.
 */
export const EDITOR_CONTRACT: BindingContract<
	XlsxEditorElement,
	EditorProps,
	EditorHandle,
	EditorEventName,
	keyof EditorEventOptions
> = {
	props: {
		locale: { forward: 'changed', fallback: 'en' },
		readOnly: { forward: 'changed', fallback: false },
		theme: { forward: 'changed', fallback: 'auto' },
		authorName: { forward: 'changed', fallback: 'Author' },
		showToolbar: { forward: 'changed', fallback: true },
		showFormulaBar: { forward: 'changed', fallback: true },
		hiddenActions: {
			forward: 'changed',
			same: sameList,
			apply: (element, value) => (element.hiddenActions = [...(value ?? [])]),
		},
		// The element keeps panels whose look is unchanged, so a fresh array per render is cheap.
		ribbonAddIns: {
			forward: 'changed',
			apply: (element, value) => (element.ribbonAddIns = value ?? []),
		},
		storage: {
			forward: 'changed',
			apply: (element, value) => (element.storage = value ?? undefined),
		},
		themeColors: {
			forward: 'changed',
			same: shallowEqual,
			apply: (element, value) => (element.themeColors = { ...value }),
		},
		// The element renames itself on File > Open; only forward a name the parent changed.
		fileName: { forward: 'defined' },
		collaboration: { forward: 'changed', fallback: null },
		customization: { forward: 'changed', fallback: null },
		workbook: { forward: 'manual' },
		bytes: { forward: 'manual' },
		src: { forward: 'manual' },
	},
	events: {
		'workbook-change': { option: 'onWorkbookChange', args: eventArgs.field('workbook') },
		'workbook-error': { option: 'onWorkbookError', args: eventArgs.field('error') },
		'selection-change': { option: 'onSelectionChange' },
		'dirty-change': { option: 'onDirtyChange', args: eventArgs.field('dirty') },
		'readonly-change': readOnlyChange,
		// Cancelable: the handler gets the event itself, so a host can call preventDefault().
		'file-command': { option: 'onFileCommand', args: eventArgs.event },
		'workbook-warning': {
			option: 'onWorkbookWarning',
			args: (event) => [[...(event.detail as { warnings: string[] }).warnings]],
		},
		'ribbon-action': { option: 'onRibbonAction' },
		'ribbon-customize': {
			option: 'onRibbonCustomize',
			args: (event) => [[...(event.detail as { hiddenActions: string[] }).hiddenActions]],
		},
		'sheet-change': { option: 'onSheetChange' },
		'office-ribbon-add-in': { option: 'onRibbonAddIn' },
		'collaboration-change': { option: 'onCollaborationChange' },
	},
	handle: {
		element: 'get',
		load: 'async',
		newWorkbook: 'call',
		save: 'async',
		saveBytes: 'async',
		download: 'async',
		markClean: 'optional-call',
		select: 'call',
		getSelection: 'call',
		setActiveSheet: 'call',
		share: 'call',
		dirty: { fallback: false },
	},
};
/** `onReady` is a callback of the binding, not an element event. */
export const EDITOR_CALLBACKS = ['onReady'] as const;

/** Copies exactly the shared props out of a framework's props/instance object. */
export function pickEditorProps(source: EditorProps): EditorProps {
	return pickProps<EditorProps>(EDITOR_CONTRACT, source);
}
/** Maps per-event handlers (keyed by DOM event name) onto the shared option callbacks. */
export function eventOptions(handlers: EditorEventHandlers): EditorEventOptions {
	return {
		...mapEventOptions<EditorEventOptions>(EDITOR_CONTRACT, handlers),
		onReady: handlers.ready,
	};
}

/** Options from Svelte's props, whose callbacks are the lower-case options (`onworkbookchange`). */
export function svelteEditorOptions(props: object): EditorOptions {
	return contractOptions<EditorOptions>(
		EDITOR_CONTRACT,
		props,
		svelteHandlerName,
		EDITOR_CALLBACKS,
	);
}

/**
 * A handle for adapters that hold a binding which may not be mounted yet: every call reads the
 * current binding and throws a clear error before mount.
 */
export function deferredHandle(current: () => EditorBinding | null | undefined): EditorHandle {
	return kitDeferredHandle(EDITOR_CONTRACT.handle, current);
}

/**
 * The imperative handle every framework exposes (React ref, Vue template ref, Angular component
 * instance, Solid `editorRef`, Svelte component instance). `contract-parity.test.ts` checks each.
 */
export const EDITOR_HANDLE_KEYS = [
	'element',
	'load',
	'newWorkbook',
	'save',
	'saveBytes',
	'download',
	'markClean',
	'select',
	'getSelection',
	'setActiveSheet',
	'share',
	'dirty',
] as const satisfies readonly (keyof EditorHandle)[];
export type EditorHandleKey = (typeof EDITOR_HANDLE_KEYS)[number];
// Compile-time guard: adding a member to EditorHandle without listing it above is an error.
const handleKeysAreComplete: Exclude<keyof EditorHandle, EditorHandleKey> extends never
	? true
	: never = true;
void handleKeysAreComplete;
