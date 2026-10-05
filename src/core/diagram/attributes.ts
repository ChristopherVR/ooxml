// Attribute-level parsers shared by every consumer of DiagramML, driven through an
// `AttributeReader` so they work on a DOM element (docx, the diagram area) and on pptx's
// `fast-xml-parser` object trees alike. Extracted from `pptx/core/utils/smartart-data-model-attributes.ts`.
import type {
	AttributeReader,
	DiagramConnection,
	DiagramNodeCustomLayout,
	DiagramRelationshipIds,
} from './types.js';

const optionalString = (value: string | undefined): string | undefined => {
	const text = (value ?? '').trim();
	return text.length > 0 ? text : undefined;
};

const optionalInteger = (value: string | undefined): number | undefined => {
	const parsed = Number.parseInt(value ?? '', 10);
	return Number.isFinite(parsed) ? parsed : undefined;
};

/** `"1"`/`"true"` -> true, anything else non-empty -> false, empty -> undefined. */
function optionalBoolean(value: string | undefined): boolean | undefined {
	const text = (value ?? '').trim();
	if (text.length === 0) return undefined;
	return text === '1' || text.toLowerCase() === 'true';
}

function optionalScaled(value: string | undefined, divisor: number): number | undefined {
	const text = (value ?? '').trim();
	if (text.length === 0) return undefined;
	const parsed = Number.parseFloat(text);
	return Number.isFinite(parsed) ? parsed / divisor : undefined;
}

/** Parses the typed CT_Cxn attributes; `undefined` without `srcId` and `destId`. */
export function parseConnectionAttributes(get: AttributeReader): DiagramConnection | undefined {
	const sourceId = optionalString(get('srcId'));
	const destId = optionalString(get('destId'));
	if (!sourceId || !destId) return undefined;
	const parsed: DiagramConnection = { sourceId, destId };
	const modelId = optionalString(get('modelId'));
	const type = optionalString(get('type'));
	const srcOrd = optionalInteger(get('srcOrd'));
	const destOrd = optionalInteger(get('destOrd'));
	const parentTransitionId = optionalString(get('parTransId'));
	const siblingTransitionId = optionalString(get('sibTransId'));
	const presentationId = optionalString(get('presId'));
	if (modelId !== undefined) parsed.modelId = modelId;
	if (type !== undefined) parsed.type = type;
	if (srcOrd !== undefined) parsed.srcOrd = srcOrd;
	if (destOrd !== undefined) parsed.destOrd = destOrd;
	if (parentTransitionId !== undefined) parsed.parentTransitionId = parentTransitionId;
	if (siblingTransitionId !== undefined) parsed.siblingTransitionId = siblingTransitionId;
	if (presentationId !== undefined) parsed.presentationId = presentationId;
	return parsed;
}

const CUSTOM_LAYOUT_ATTRIBUTES: ReadonlyArray<
	[keyof DiagramNodeCustomLayout, string, 'angle' | 'ratio' | 'flag']
> = [
	['angle', 'custAng', 'angle'],
	['scaleX', 'custScaleX', 'ratio'],
	['scaleY', 'custScaleY', 'ratio'],
	['sizeX', 'custSzX', 'ratio'],
	['sizeY', 'custSzY', 'ratio'],
	['linearFactorX', 'custLinFactX', 'ratio'],
	['linearFactorY', 'custLinFactY', 'ratio'],
	['linearFactorNeighborX', 'custLinFactNeighborX', 'ratio'],
	['linearFactorNeighborY', 'custLinFactNeighborY', 'ratio'],
	['radialScaleRadius', 'custRadScaleRad', 'ratio'],
	['radialScaleIncrement', 'custRadScaleInc', 'ratio'],
	['flipHorizontal', 'custFlipHor', 'flag'],
	['flipVertical', 'custFlipVert', 'flag'],
	['hasCustomTransform', 'custT', 'flag'],
];

/**
 * Parses the manual layout override attributes (`cust*`) of a `dgm:prSet`, or `undefined` when it
 * carries none. Angles are 60,000ths of a degree on disk and plain degrees here; scale and size
 * factors are 100,000ths on disk and plain ratios here.
 */
export function parseCustomLayoutAttributes(
	get: AttributeReader,
): DiagramNodeCustomLayout | undefined {
	const custom: Record<string, number | boolean> = {};
	for (const [key, attribute, kind] of CUSTOM_LAYOUT_ATTRIBUTES) {
		const raw = get(attribute);
		const value =
			kind === 'flag'
				? optionalBoolean(raw)
				: optionalScaled(raw, kind === 'angle' ? 60000 : 100000);
		if (value !== undefined) custom[key] = value;
	}
	return Object.keys(custom).length > 0 ? (custom as DiagramNodeCustomLayout) : undefined;
}

const RELATIONSHIP_ATTRIBUTES: ReadonlyArray<[keyof DiagramRelationshipIds, string]> = [
	['dataRelId', 'dm'],
	['layoutRelId', 'lo'],
	['styleRelId', 'qs'],
	['colorsRelId', 'cs'],
];

/** Reads `CT_RelIds` (`r:dm`, `r:lo`, `r:qs`, `r:cs`); the reader matches attributes by local name. */
export function parseRelationshipIdAttributes(get: AttributeReader): DiagramRelationshipIds {
	const result: DiagramRelationshipIds = {};
	for (const [property, attribute] of RELATIONSHIP_ATTRIBUTES) {
		const value = optionalString(get(attribute));
		if (value) result[property] = value;
	}
	return result;
}
