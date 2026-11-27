/**
 * The picture half of a fabricated SmartArt (Picture Format > Picture Layout): the new diagram's
 * picture placeholders take the converted pictures.
 *
 * The in-memory diagram names each picture by its node id and a package part; a picture that was
 * edited or inserted in this session has no part yet and carries a pending key with its data URL.
 * Saving stores those bytes as media, maps node ids to the data model's model ids, and lets the
 * format-neutral cache writer lay the diagram out with the pictures in its placeholders, so the
 * data part, the drawing cache and both parts' image relationships agree.
 *
 * @module runtime/smartart-fabrication-pictures
 */
import type { DiagramLayoutSource } from '../../../../diagram/layout-preview';
import {
	diagramPictureRelationships,
	withDiagramPictureSources,
} from '../../../../diagram/picture-layout';
import {
	buildGeneratedDiagramCache,
	withDiagramDrawingRelationship,
} from '../../../../diagram/write/cache';
import type { PptxSmartArtData } from '../../types';

export interface FabricatedPictureInput {
	data: PptxSmartArtData;
	dataXml: string;
	layoutXml: string;
	dataPath: string;
	drawingPath: string;
	drawingRelId: string | undefined;
	guidByNodeId: ReadonlyMap<string, string>;
	extentEmu: { width: number; height: number };
	createModelId: () => string;
	/** Whether the package already holds `part`. */
	hasPart: (part: string) => boolean;
	/** Store a data URL's bytes as a media part; its path, or `undefined` when not a data URL. */
	storeDataUrl: (url: string) => string | undefined;
}

export interface FabricatedPictureParts {
	dataXml: string;
	drawingXml?: string | undefined;
	dataRelsXml?: string | undefined;
	drawingRelsXml?: string | undefined;
}

/** The picture source with saved parts and model ids, or `undefined` when no picture survives. */
function savedPictureSource(input: FabricatedPictureInput): DiagramLayoutSource | undefined {
	const source = input.data.pictureLayoutSource;
	if (!source?.pictures?.length || !source.pictureParts) return undefined;
	const parts: Record<string, string> = {};
	for (const [relId, part] of Object.entries(source.pictureParts)) {
		const url = input.data.pictureUrls?.[part];
		const stored = url ? input.storeDataUrl(url) : undefined;
		if (stored) parts[relId] = stored;
		else if (input.hasPart(part)) parts[relId] = part;
	}
	const pictures = source.pictures.flatMap((picture) => {
		const nodeId = input.guidByNodeId.get(picture.nodeId);
		const relId = picture.fill.relId;
		return nodeId && relId && parts[relId] ? [{ ...picture, nodeId }] : [];
	});
	return pictures.length
		? { ...source, pictures, pictureParts: parts, layoutXml: input.layoutXml }
		: undefined;
}

/** Lay out a fabricated picture diagram; `undefined` leaves the plain fabricated parts. */
export function fabricatePictureLayoutParts(
	input: FabricatedPictureInput,
): FabricatedPictureParts | undefined {
	const source = savedPictureSource(input);
	if (!source) return undefined;
	const cache = input.drawingRelId
		? buildGeneratedDiagramCache(
				input.dataXml,
				{
					extentEmu: input.extentEmu,
					source,
					minorFont: input.data.themeMinorFont,
					scheme: input.data.pictureScheme,
				},
				input.createModelId,
			)
		: undefined;
	const dataRelsXml = diagramPictureRelationships(input.dataPath, source);
	if (!cache || !input.drawingRelId) {
		// No generated cache: the associations still persist, so a reload or PowerPoint's own
		// layout pass puts the pictures back into the placeholders.
		return { dataXml: withDiagramPictureSources(input.dataXml, source), dataRelsXml };
	}
	return {
		dataXml: withDiagramDrawingRelationship(cache.dataXml, input.drawingRelId),
		drawingXml: cache.drawingXml,
		dataRelsXml,
		drawingRelsXml: diagramPictureRelationships(input.drawingPath, source),
	};
}
