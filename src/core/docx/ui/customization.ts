/**
 * Word's half of the shared customization contract (`ooxml-core/embedding`): the capabilities
 * `<docx-editor>` honours and what each one removes. DOM-free, so a host or a binding can read
 * it without loading the editor.
 */
import type {
	CapabilityActionMap,
	EditorCapabilities,
	EditorCapabilityId,
} from '../../embedding/index';
import type { RibbonActionId } from './ribbon-action-ids';

/** The capabilities Word honours. */
export type WordCapabilityId = Extract<
	EditorCapabilityId,
	| 'edit'
	| 'export'
	| 'print'
	| 'share'
	| 'copy'
	| 'comment'
	| 'review'
	| 'find'
	| 'open'
	| 'save'
	| 'settings'
>;

/** Every Word capability, for hosts that build a policy screen. */
export const WORD_CAPABILITY_IDS = [
	'edit',
	'export',
	'print',
	'share',
	'copy',
	'comment',
	'review',
	'find',
	'open',
	'save',
	'settings',
] as const satisfies readonly WordCapabilityId[];

/** Per-capability switches of `<docx-editor>`. */
export type WordCapabilities = EditorCapabilities<WordCapabilityId>;

/** The ribbon controls (`hiddenActions` ids) each disabled capability removes. */
export const WORD_CAPABILITY_ACTIONS: CapabilityActionMap<WordCapabilityId, RibbonActionId> = {
	print: ['print'],
	copy: ['cut', 'copy'],
	comment: ['add-comment', 'delete-comment', 'previous-comment', 'next-comment', 'comments'],
	review: ['track-changes', 'tracking-options', 'accept', 'reject', 'accept-all', 'reject-all'],
	find: ['find-and-replace', 'replace', 'find-options'],
};

/** A File command (title bar, File tab, shortcut) that a capability can remove. */
export type WordFileCommandId =
	| 'new'
	| 'open'
	| 'save'
	| 'saveAs'
	| 'export'
	| 'exportText'
	| 'print'
	| 'share';

/** The File commands each disabled capability removes (and refuses to run). */
export const WORD_CAPABILITY_FILE_COMMANDS: CapabilityActionMap<
	WordCapabilityId,
	WordFileCommandId
> = {
	open: ['new', 'open'],
	save: ['save'],
	export: ['saveAs', 'export', 'exportText'],
	print: ['print'],
	share: ['share'],
};

/** A File tab navigation entry a capability can remove. */
export type WordBackstagePageId =
	| 'new'
	| 'open'
	| 'save'
	| 'saveAs'
	| 'print'
	| 'export'
	| 'options';

/** The File tab entries each disabled capability removes. */
export const WORD_CAPABILITY_BACKSTAGE_PAGES: CapabilityActionMap<
	WordCapabilityId,
	WordBackstagePageId
> = {
	open: ['new', 'open'],
	save: ['save'],
	export: ['saveAs', 'export'],
	print: ['print'],
	settings: ['options'],
};

/** A built-in keyboard shortcut a capability can turn off. */
export type WordShortcutId = 'find' | 'replace' | 'save' | 'print';

/** The shortcuts each disabled capability turns off. */
export const WORD_CAPABILITY_SHORTCUTS: CapabilityActionMap<WordCapabilityId, WordShortcutId> = {
	find: ['find', 'replace'],
	save: ['save'],
	print: ['print'],
};

/** The built-in context-menu entries each disabled capability removes. */
export const WORD_CAPABILITY_MENU_ITEMS: CapabilityActionMap<WordCapabilityId, string> = {
	copy: ['cut', 'copy'],
	comment: ['comment-add'],
};
