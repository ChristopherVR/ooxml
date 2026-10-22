/**
 * The shared shape of an editor's `customization` object. A host describes what its users see
 * and may do in ONE plain object; every product (Word, Excel, PowerPoint, Visio) accepts the
 * same sections through its element property and every framework binding forwards it as one
 * prop. Every field is optional and every list defaults to "hide nothing", so an empty object
 * (or none) is exactly the editor's default behaviour.
 *
 * Products bind the type parameters: `Capability` to the ids they honour, `Renderer` to their
 * region renderer, `MenuContext` to the snapshot their context menu passes to host commands.
 * PowerPoint's `ViewerCustomization` is a superset of this shape.
 */
import type { HostBackstagePage } from './backstage';
import type { EditorCapabilities } from './capabilities';
import type { HostMenuCommand } from './menu-commands';

/** File tab (Backstage) customisation every product supports. */
export interface EditorBackstageCustomization<Renderer> {
	/** Host pages added after the built-in pages of their group. */
	extraPages?: readonly HostBackstagePage<Renderer>[];
}

/** Right-click menu customisation every product supports. */
export interface EditorContextMenuCustomization<MenuContext> {
	/** Host commands added to the document (cell, shape) context menu. */
	extraCommands?: readonly HostMenuCommand<MenuContext>[];
}

/** The customisation sections every editor shares. */
export interface EditorCustomization<Capability extends string, Renderer, MenuContext> {
	/** What the user may do; `false` removes every UI path to the capability. */
	capabilities?: EditorCapabilities<Capability>;
	backstage?: EditorBackstageCustomization<Renderer>;
	contextMenu?: EditorContextMenuCustomization<MenuContext>;
}
