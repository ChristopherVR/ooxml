import type { CanvasContextMenuEntry } from '../canvas-context-menu-commands';
import type { ContextMenuEntry } from '../context-menu-commands';
import {
	withHostMenuCommands,
	type HostMenuCommand,
	type HostMenuEntry,
} from 'ooxml-core/embedding';

/** Snapshot of the slide on which the host command was offered. */
export interface HostCanvasMenuContext {
	slideIndex: number;
}

/** Effective selection, including a single right-clicked element. */
export interface HostElementMenuContext extends HostCanvasMenuContext {
	elementIds: readonly string[];
}

/**
 * Host commands and the menu entries made from them are the shared ones of
 * `ooxml-core/embedding`; PowerPoint passes the slide and selection as context.
 */
export { withHostMenuCommands, type HostMenuCommand, type HostMenuEntry };

export type CustomizedContextMenuEntry = ContextMenuEntry | HostMenuEntry;
export type CustomizedCanvasContextMenuEntry = CanvasContextMenuEntry | HostMenuEntry;

/** Host labels are already localized by the host. */
export function hostMenuLabel(
	entry: CustomizedContextMenuEntry | CustomizedCanvasContextMenuEntry,
): string | undefined {
	return 'host' in entry ? entry.label : undefined;
}
