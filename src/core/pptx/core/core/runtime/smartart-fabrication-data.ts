/**
 * Fabricate the diagram DATA part (`ppt/diagrams/dataN.xml`) for an
 * SDK-created SmartArt element. The writer lives in `diagram/write/data`; this
 * facade keeps the pptx names and mints model ids through the pptx GUID
 * generator.
 */
import {
	buildDiagramNodeGuidMap,
	buildDiagramPresentationGuidMap,
	writeFabricatedDiagramData,
	type DiagramDefinitionIds,
	type FabricatedDiagramDataOptions,
} from '../../../../diagram/write/data';
import type { PptxSmartArtData } from '../../types';
import { newSmartArtGuid } from './smartart-xml-builders';

export {
	DIAGRAM_DATA_XMLNS as DGM_XMLNS,
	DIAGRAM_XML_PROLOG as XML_PROLOG,
	diagramAvLstXml as avLstXml,
	escapeDiagramXml as xmlEscape,
	type FabricatedDiagramDataOptions,
} from '../../../../diagram/write/data';

/** Unique-id URNs of the fabricated layout / quick-style / colors parts. */
export type FabricatedDefIds = DiagramDefinitionIds;

/** The stable `nodeId -> {GUID}` model-id map shared by the data and drawing parts. */
export function buildNodeGuidMap(nodes: PptxSmartArtData['nodes']): Map<string, string> {
	return buildDiagramNodeGuidMap(nodes, newSmartArtGuid);
}

/** Stable presentation-point ids for cached drawing shapes. */
export function buildPresentationGuidMap(nodes: PptxSmartArtData['nodes']): Map<string, string> {
	return buildDiagramPresentationGuidMap(nodes, newSmartArtGuid);
}

/** Build the complete `dataN.xml` payload for a fabricated SmartArt diagram. */
export function buildFabricatedDiagramDataXml(
	data: PptxSmartArtData,
	ids: DiagramDefinitionIds,
	options: FabricatedDiagramDataOptions = {},
): string {
	return writeFabricatedDiagramData(data, ids, { createModelId: newSmartArtGuid, ...options });
}
