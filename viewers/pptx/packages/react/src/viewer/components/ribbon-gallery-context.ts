/**
 * The ribbon style galleries' link to the viewer: the gallery context the
 * shared descriptors are built from (selection, theme, style-matrix resolver)
 * and the dispatcher that turns a pick's apply result into an undoable edit.
 *
 * Provided once by `PowerPointViewer` (the same reason as
 * `ShapeFormatContext`: the galleries sit in the desktop ribbon, the
 * contextual tabs and the Design tab, and threading this through every
 * section's props would touch half a dozen oversized components). Absent a
 * provider, as in a standalone unit test, galleries build from an empty
 * context and every trigger renders disabled.
 */
import type { RibbonGalleryApplyResult, RibbonGalleryContext } from 'ooxml-ui/pptx';
import { createContext, useContext } from 'react';

export interface RibbonGalleryCommands {
	/** What `buildRibbonGallery` / `applyRibbonGalleryItem` read. */
	context: RibbonGalleryContext;
	/** False on a read-only deck: every gallery trigger disables. */
	editable: boolean;
	/** Route a pick's result onto the viewer's update + history paths. */
	dispatch: (result: RibbonGalleryApplyResult) => void;
	/** Show what a Design gallery tile would do while it is hovered; `null` ends the preview. */
	preview?: (result: RibbonGalleryApplyResult | null) => void;
	/** Put the slides back (the pointer left the tiles, or the gallery closed). */
	endPreview?: () => void;
}

export const RibbonGalleryCommandsContext =
	/* @__PURE__ */ createContext<RibbonGalleryCommands | null>(null);

export function useRibbonGalleryCommands(): RibbonGalleryCommands | null {
	return useContext(RibbonGalleryCommandsContext);
}
