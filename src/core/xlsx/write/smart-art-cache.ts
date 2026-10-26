import {
	buildGeneratedDiagramCache,
	withDiagramDrawingRelationship,
	DIAGRAM_RELATIONSHIP_TYPES,
	DIAGRAM_CONTENT_TYPES,
} from '../../diagram/index';
import { drawingSchemeColors } from '../drawing-theme';
import type { SmartArtObject, ThemePalette } from '../model';
import type { PackageWriter, RelationshipSet } from './package-writer';

export interface SmartArtCacheContext {
	theme?: ThemePalette;
	extentEmu?: { width: number; height: number };
}

/** Attach an admitted generated cache to the worksheet drawing's relationship set. */
export function writeSmartArtCache(
	writer: PackageWriter,
	object: SmartArtObject,
	parts: { data: string; layout?: string; colors?: string; quickStyle?: string },
	rels: RelationshipSet,
	relativeTarget: (part: string) => string,
	context: SmartArtCacheContext,
): void {
	const text = (name: string | undefined): string | undefined => {
		if (!name) return undefined;
		const pending = writer.peek(name);
		return pending instanceof Uint8Array
			? new TextDecoder().decode(pending)
			: (pending ?? writer.source?.text(name));
	};
	const data = text(parts.data);
	if (!data) return;
	const layoutXml = text(parts.layout);
	const colorsXml = text(parts.colors);
	const quickStyleXml = text(parts.quickStyle);
	const cache = buildGeneratedDiagramCache(data, {
		extentEmu: context.extentEmu ?? {
			width: object.anchor.ext?.cx ?? 4572000,
			height: object.anchor.ext?.cy ?? 2667000,
		},
		source: {
			...(layoutXml ? { layoutXml } : {}),
			...(colorsXml ? { colorsXml } : {}),
			...(quickStyleXml ? { quickStyleXml } : {}),
		},
		layoutId: object.layoutId,
		family: object.layoutFamily,
		scheme: drawingSchemeColors(context.theme),
		minorFont: context.theme?.minorFont,
		formatScheme: context.theme?.formatScheme,
	});
	if (!cache) return;
	let index = 1;
	while (
		writer.has(`xl/diagrams/drawing${index}.xml`) ||
		writer.source?.has(`xl/diagrams/drawing${index}.xml`)
	)
		index++;
	const drawingName = `xl/diagrams/drawing${index}.xml`;
	const relationship = rels.add(DIAGRAM_RELATIONSHIP_TYPES.drawing, relativeTarget(drawingName));
	writer.add(drawingName, cache.drawingXml, DIAGRAM_CONTENT_TYPES.drawing);
	writer.add(
		parts.data,
		withDiagramDrawingRelationship(cache.dataXml, relationship),
		DIAGRAM_CONTENT_TYPES.data,
	);
}
