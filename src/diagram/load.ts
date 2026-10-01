// Loads every part of one diagram through a host-supplied part reader, reporting what is missing
// or unreadable instead of throwing, so a host can show an honest fallback.
import {
	parseDiagramColors,
	parseDiagramLayoutSummary,
	parseDiagramQuickStyle,
} from './definitions.js';
import { parseDiagramDataModel } from './data-model.js';
import { parseDiagramDrawing } from './drawing.js';
import { resolveDiagramDrawingPart, type DiagramPartHost } from './relationships.js';
import type {
	DiagramColorsDefinition,
	DiagramDataModel,
	DiagramDrawing,
	DiagramIssue,
	DiagramLayoutSummary,
	DiagramQuickStyleDefinition,
	DiagramRelationshipIds,
} from './types.js';

/** A part host that can also read part text. */
export interface DiagramPackageHost extends DiagramPartHost {
	/** The text of a part, or `undefined` when the package has no such part. */
	readText(partPath: string): Promise<string | undefined>;
}

/** Package paths of the diagram parts that could be resolved. */
export interface DiagramPartPaths {
	data?: string;
	layout?: string;
	quickStyle?: string;
	colors?: string;
	drawing?: string;
}

/** Every part of a diagram, parsed, with the problems met on the way. */
export interface LoadedDiagram {
	relationshipIds: DiagramRelationshipIds;
	paths: DiagramPartPaths;
	/** Relationship id (in the host part) of the cached drawing, when one was found. */
	drawingRelId?: string;
	data?: DiagramDataModel;
	layout?: DiagramLayoutSummary;
	colors?: DiagramColorsDefinition;
	quickStyle?: DiagramQuickStyleDefinition;
	/** The cached layout. Absent when the producer wrote none: only the data model then describes the diagram. */
	drawing?: DiagramDrawing;
	issues: DiagramIssue[];
}

const slot = [
	['data', 'dataRelId'],
	['layout', 'layoutRelId'],
	['quickStyle', 'styleRelId'],
	['colors', 'colorsRelId'],
] as const;

/**
 * Resolves and parses the five parts of a diagram referenced by `relationshipIds` from
 * `hostPart` (a slide, `word/document.xml`, a header...). Missing relationships, missing parts and
 * parse failures become `issues`; the other parts still load.
 */
export async function loadDiagram(
	hostPart: string,
	relationshipIds: DiagramRelationshipIds,
	host: DiagramPackageHost,
): Promise<LoadedDiagram> {
	const issues: DiagramIssue[] = [];
	const paths: DiagramPartPaths = {};
	const hostRels = host.hostRelationships(hostPart);
	for (const [name, key] of slot) {
		const relId = relationshipIds[key];
		if (!relId) {
			issues.push({
				code: 'DIAGRAM_RELATIONSHIP_MISSING',
				message: `The diagram has no ${name} relationship id.`,
			});
			continue;
		}
		const target = hostRels?.get(relId);
		if (target === undefined)
			issues.push({
				code: 'DIAGRAM_RELATIONSHIP_UNRESOLVED',
				message: `Relationship ${relId} (${name}) is not in ${hostPart}.`,
			});
		else paths[name] = host.resolvePath(hostPart, target);
	}

	const read = async <T>(
		name: keyof DiagramPartPaths,
		parse: (xml: string, part: string) => T,
	): Promise<T | undefined> => {
		const part = paths[name];
		if (!part) return undefined;
		try {
			const xml = await host.readText(part);
			if (xml === undefined) {
				issues.push({
					code: 'DIAGRAM_PART_MISSING',
					message: `Part ${part} is not in the package.`,
					part,
				});
				return undefined;
			}
			return parse(xml, part);
		} catch (error) {
			issues.push({
				code: 'DIAGRAM_PART_UNREADABLE',
				message: `Part ${part} could not be parsed: ${error instanceof Error ? error.message : String(error)}`,
				part,
			});
			return undefined;
		}
	};

	const data = await read('data', (xml, part) => parseDiagramDataModel(xml, part));
	if (data) issues.push(...data.issues);
	const layout = await read('layout', (xml) => parseDiagramLayoutSummary(xml));
	const colors = await read('colors', (xml) => parseDiagramColors(xml));
	const quickStyle = await read('quickStyle', (xml) => parseDiagramQuickStyle(xml));

	let drawingResolution: { relId: string; path: string } | undefined;
	try {
		drawingResolution = await resolveDiagramDrawingPart(
			hostPart,
			relationshipIds.dataRelId ?? '',
			data?.drawingRelId ?? '',
			host,
		);
	} catch (error) {
		issues.push({
			code: 'DIAGRAM_DRAWING_UNRESOLVED',
			message: `The cached drawing could not be located: ${error instanceof Error ? error.message : String(error)}`,
		});
	}
	let drawing: DiagramDrawing | undefined;
	if (drawingResolution) {
		paths.drawing = drawingResolution.path;
		drawing = await read('drawing', (xml, part) => parseDiagramDrawing(xml, part));
		if (drawing) issues.push(...drawing.issues);
	} else if (data)
		issues.push({
			code: 'DIAGRAM_DRAWING_ABSENT',
			message:
				'The package has no cached drawing for this diagram; its layout can only be recomputed from the layout definition.',
		});

	return {
		relationshipIds,
		paths,
		...(drawingResolution ? { drawingRelId: drawingResolution.relId } : {}),
		...(data ? { data } : {}),
		...(layout ? { layout } : {}),
		...(colors ? { colors } : {}),
		...(quickStyle ? { quickStyle } : {}),
		...(drawing ? { drawing } : {}),
		issues,
	};
}
