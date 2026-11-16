/**
 * Saving a built-in colour or quick-style definition picked in the SmartArt Design galleries:
 * the definition replaces the part (PowerPoint writes the whole definition of the style it
 * applies), and the document point's `csTypeId`/`csCatId` or `qsTypeId`/`qsCatId` name it.
 */
import type { PptxSmartArtData, XmlObject } from '../../types';

const PROLOG = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n';

export interface SmartArtDefinitionIdentity {
	uniqueId: string;
	category?: string;
}

/** The `uniqueId` and first category of a definition's XML. */
export function smartArtDefinitionIdentity(xml: string): SmartArtDefinitionIdentity | undefined {
	const uniqueId = /<[\w:]*(?:colorsDef|styleDef)\b[^>]*\buniqueId="([^"]+)"/u.exec(xml)?.[1];
	if (!uniqueId) return undefined;
	const category = /<[\w:]*cat\b[^>]*\btype="([^"]+)"/u.exec(xml)?.[1];
	return category ? { uniqueId, category } : { uniqueId };
}

/** The part content for a built-in definition. */
export const smartArtDefinitionPartXml = (xml: string): string =>
	xml.startsWith('<?xml') ? xml : `${PROLOG}${xml}`;

/** Name the built-in colour and quick-style definitions on the document point's `dgm:prSet`. */
export function applySmartArtDefinitionIdentity(
	dataModel: XmlObject,
	data: Pick<PptxSmartArtData, 'builtinColorsXml' | 'builtinQuickStyleXml'>,
	localName: (key: string) => string,
): void {
	const colors = data.builtinColorsXml
		? smartArtDefinitionIdentity(data.builtinColorsXml)
		: undefined;
	const style = data.builtinQuickStyleXml
		? smartArtDefinitionIdentity(data.builtinQuickStyleXml)
		: undefined;
	if (!colors && !style) return;
	const pointListKey = Object.keys(dataModel).find((key) => localName(key) === 'ptLst');
	const pointList = pointListKey ? (dataModel[pointListKey] as XmlObject) : undefined;
	const pointKey = pointList
		? Object.keys(pointList).find((key) => localName(key) === 'pt')
		: undefined;
	const raw = pointKey && pointList ? pointList[pointKey] : undefined;
	const points = (Array.isArray(raw) ? raw : raw ? [raw] : []) as XmlObject[];
	const docPoint = points.find((point) => String(point['@_type'] || '') === 'doc');
	const propertySetKey = docPoint
		? Object.keys(docPoint).find((key) => localName(key) === 'prSet')
		: undefined;
	const propertySet =
		docPoint && propertySetKey ? (docPoint[propertySetKey] as XmlObject) : undefined;
	if (!propertySet || typeof propertySet !== 'object') return;
	if (colors) {
		propertySet['@_csTypeId'] = colors.uniqueId;
		if (colors.category) propertySet['@_csCatId'] = colors.category;
	}
	if (style) {
		propertySet['@_qsTypeId'] = style.uniqueId;
		if (style.category) propertySet['@_qsCatId'] = style.category;
	}
}
