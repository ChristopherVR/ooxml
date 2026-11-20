/**
 * The shared shape of an editor's `customization` object. A host describes what its users see
 * and may do in ONE plain object; every product (Word, Excel, PowerPoint, Visio) accepts the
 * same sections through its element property and every framework binding forwards it as one
 * prop. Every field is optional and every list defaults to "hide nothing", so an empty object
 * (or none) is exactly the editor's default behaviour.
 *
 * Products bind the type parameters: `Capability` to the ids they honour, `Renderer` to their
 * region renderer, `MenuContext` to the snapshot their context menu passes to host commands and
 * `Tab` to their native ribbon tab ids. PowerPoint's `ViewerCustomization` is a superset of this
 * shape.
 */
import type { HostBackstagePage } from './backstage';
import type { EditorCapabilities } from './capabilities';
import type { HostMenuCommand } from './menu-commands';

/** The id every product gives its File tab (and the backstage behind it) in `ribbon.hiddenTabs`. */
export const FILE_RIBBON_TAB = 'file';

/** File tab (Backstage) customisation every product supports. */
export interface EditorBackstageCustomization<Renderer> {
	/** Host pages added after the built-in pages of their group. */
	extraPages?: readonly HostBackstagePage<Renderer>[];
}

/** Ribbon customisation every product supports. */
export interface EditorRibbonCustomization<Tab extends string = typeof FILE_RIBBON_TAB> {
	/**
	 * Native ribbon tabs to remove, contextual tabs included. `file` removes the File tab and the
	 * backstage behind it: no key tip, shortcut, status-bar entry or command opens it, while Save,
	 * Ctrl+S, `storage` and the `file-command` event keep working. Host tabs (`ribbonAddIns`) are
	 * not affected. Unknown ids are ignored.
	 */
	hiddenTabs?: readonly Tab[];
}

/** The customisation sections every editor shares. */
export interface EditorCustomization<
	Capability extends string,
	Renderer,
	MenuContext,
	Tab extends string = typeof FILE_RIBBON_TAB,
> {
	/** What the user may do; `false` removes every UI path to the capability. */
	capabilities?: EditorCapabilities<Capability>;
	ribbon?: EditorRibbonCustomization<Tab>;
	backstage?: EditorBackstageCustomization<Renderer>;
	contextMenu?: EditorContextMenuCustomization<MenuContext>;
}

/** Right-click menu customisation every product supports. */
export interface EditorContextMenuCustomization<MenuContext> {
	/** Host commands added to the document (cell, shape) context menu. */
	extraCommands?: readonly HostMenuCommand<MenuContext>[];
}

/** The tabs a host removed (`ribbon.hiddenTabs`), deduplicated; empty when none. */
export function hiddenRibbonTabs<Tab extends string>(
	ribbon: EditorRibbonCustomization<Tab> | null | undefined,
): ReadonlySet<Tab> {
	return new Set(ribbon?.hiddenTabs ?? []);
}
