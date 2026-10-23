// Contracts of the format-neutral SmartArt edit operations (`diagram/edit`). A product's parsed
// SmartArt (pptx: `PptxSmartArtData`) satisfies `DiagramEditableData` structurally, so the
// operations return the product's own type unchanged apart from the edited fields.
import type { DiagramConnection } from '../types';
import type { DiagramNode } from '../model/node';

/** A drawing shape as the node operations see it: the id they match and the text they patch. */
export interface DiagramEditShape {
	id: string;
	text?: string | undefined;
}

/** The parts of a parsed SmartArt graphic the node operations read and replace. */
export interface DiagramEditableData<
	N extends DiagramNode = DiagramNode,
	S extends DiagramEditShape = DiagramEditShape,
> {
	nodes: N[];
	connections?: DiagramConnection[] | undefined;
	/** Cached or reflowed drawing shapes; `[]` after an edit marks them stale. */
	drawingShapes?: S[] | undefined;
	/** Save hint: the cached drawing no longer matches the data model. */
	drawingDirty?: boolean | undefined;
}

export interface DiagramEditOptions {
	/**
	 * Mints the `modelId` of a new point. Defaults to a braced upper-case GUID, the form PowerPoint
	 * requires (`{XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX}`); PowerPoint rejects other ids as corrupt.
	 */
	createModelId?: () => string;
}

/** The rectangle a reflow lays shapes out in, in the host's units (pixels for pptx). */
export interface DiagramReflowBounds {
	x: number;
	y: number;
	width: number;
	height: number;
}
