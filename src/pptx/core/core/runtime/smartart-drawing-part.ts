import {
	resolveDiagramDrawingPart,
	type DiagramRelationshipEntry,
} from '../../../../diagram/index.js';
import type { XmlObject } from '../../types';

interface DrawingPartDeps {
	slideRelationships(path: string): Map<string, string> | undefined;
	resolvePath(base: string, target: string): string;
	readText(path: string): Promise<string | undefined>;
	parse(xml: string): XmlObject;
	ensureArray(value: unknown): unknown[];
}

/**
 * Resolve slide-scoped diagram drawings, with a legacy data-part relationship fallback. The
 * resolution order lives in the format-neutral `diagram` area; this adapter supplies the slide's
 * relationship lookup and reads `.rels` parts through the presentation's own XML parser.
 */
export function resolveSmartArtDrawingPart(
	slidePath: string,
	diagramDataRelationshipId: string,
	drawingExtensionRelId: string,
	deps: DrawingPartDeps,
): Promise<{ relId: string; path: string } | undefined> {
	return resolveDiagramDrawingPart(slidePath, diagramDataRelationshipId, drawingExtensionRelId, {
		hostRelationships: (path) => deps.slideRelationships(path),
		resolvePath: (base, target) => deps.resolvePath(base, target),
		readRelationshipEntries: async (relsPath): Promise<DiagramRelationshipEntry[] | undefined> => {
			const relsXml = await deps.readText(relsPath);
			if (!relsXml) {
				return undefined;
			}
			try {
				const parsed = deps.parse(relsXml) as XmlObject;
				const relsRoot = parsed['Relationships'] as XmlObject | undefined;
				if (!relsRoot) {
					return undefined;
				}
				return (deps.ensureArray(relsRoot['Relationship']) as XmlObject[]).map((rel) => ({
					id: String(rel?.['@_Id'] || '').trim(),
					type: String(rel?.['@_Type'] || ''),
					target: String(rel?.['@_Target'] || ''),
				}));
			} catch {
				return undefined;
			}
		},
	});
}
