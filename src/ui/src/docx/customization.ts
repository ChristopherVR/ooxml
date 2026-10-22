/**
 * `<docx-editor>`'s `customization` property: the shared contract of `ooxml-core/embedding`
 * (capabilities, host File tab pages, host context-menu commands) bound to Word. The host passes
 * one plain object; `resolveWordCustomization` turns it into the sets every render site asks,
 * so the ribbon, the File tab, the context menu, the shortcuts and the file commands agree.
 */
import {
	capabilityHiddenActions,
	isCapabilityEnabled,
	type CapabilityActionMap,
	type EditorCustomization,
	type HostBackstagePage,
	type HostMenuCommand,
	type HostRegionContextBase,
	type HostRegionRenderer,
} from 'ooxml-core/embedding';
import {
	WORD_CAPABILITY_ACTIONS,
	WORD_CAPABILITY_BACKSTAGE_PAGES,
	WORD_CAPABILITY_FILE_COMMANDS,
	WORD_CAPABILITY_MENU_ITEMS,
	WORD_CAPABILITY_SHORTCUTS,
	type RibbonActionId,
	type WordBackstagePageId,
	type WordCapabilityId,
	type WordFileCommandId,
	type WordShortcutId,
} from 'ooxml-core/docx/ui';

/** What a host region (a File tab page) of `<docx-editor>` receives. */
export interface WordRegionContext extends HostRegionContextBase {
	/** The `<docx-editor>` element: its whole public API (save, documentModel, readOnly...). */
	readonly editor: HTMLElement;
	/** The document's file name. */
	readonly fileName: string;
}

/** What a host context-menu command of `<docx-editor>` receives: the menu's selection. */
export interface WordMenuContext {
	/** The selected text; empty for a caret. */
	readonly selectedText: string;
	/** Selection start and end, as document positions (equal for a caret). */
	readonly from: number;
	readonly to: number;
	/** True when the document cannot be edited (prop, capability, protection, collaboration). */
	readonly readOnly: boolean;
	readonly inTable: boolean;
	readonly onLink: boolean;
}

export type WordRegionRenderer = HostRegionRenderer<HTMLElement, WordRegionContext>;
export type WordBackstagePage = HostBackstagePage<WordRegionRenderer>;
export type WordMenuCommand = HostMenuCommand<WordMenuContext>;

/** The whole `customization` object of `<docx-editor>`. Every field is optional. */
export type WordCustomization = EditorCustomization<
	WordCapabilityId,
	WordRegionRenderer,
	WordMenuContext
>;

/** A customisation resolved into what each render site asks. */
export interface ResolvedWordCustomization {
	/** False when `capabilities.edit` is off: the document opens read-only. */
	readonly canEdit: boolean;
	/** Ribbon controls the capabilities remove, unioned with `hiddenActions` by the ribbon. */
	readonly hiddenActions: readonly RibbonActionId[];
	readonly fileCommands: ReadonlySet<WordFileCommandId>;
	readonly backstagePages: ReadonlySet<WordBackstagePageId>;
	readonly shortcuts: ReadonlySet<WordShortcutId>;
	/** Built-in context-menu entries to remove. */
	readonly menuItems: ReadonlySet<string>;
	/** True when `copy` is off: the editor also refuses copy and cut from the keyboard. */
	readonly blockCopy: boolean;
	readonly extraPages: readonly WordBackstagePage[];
	readonly menuCommands: readonly WordMenuCommand[];
}

/** Resolves a host customisation (none: everything on, nothing added). */
export function resolveWordCustomization(
	customization?: WordCustomization | null,
): ResolvedWordCustomization {
	const capabilities = customization?.capabilities;
	const set = <T extends string>(map: CapabilityActionMap<WordCapabilityId, T>) =>
		new Set(capabilityHiddenActions(capabilities, map));
	return {
		canEdit: isCapabilityEnabled(capabilities, 'edit'),
		hiddenActions: capabilityHiddenActions(capabilities, WORD_CAPABILITY_ACTIONS),
		fileCommands: set(WORD_CAPABILITY_FILE_COMMANDS),
		backstagePages: set(WORD_CAPABILITY_BACKSTAGE_PAGES),
		shortcuts: set(WORD_CAPABILITY_SHORTCUTS),
		menuItems: set(WORD_CAPABILITY_MENU_ITEMS),
		blockCopy: !isCapabilityEnabled(capabilities, 'copy'),
		extraPages: customization?.backstage?.extraPages ?? [],
		menuCommands: customization?.contextMenu?.extraCommands ?? [],
	};
}

/** The resolution of "no customisation", shared so it is allocated once. */
export const EMPTY_WORD_CUSTOMIZATION: ResolvedWordCustomization = resolveWordCustomization();
