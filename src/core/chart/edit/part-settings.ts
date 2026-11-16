// Settings of a chart part outside its plot area (ECMA-376 Part 1, 21.2.2.27 CT_ChartSpace): the
// external data auto-update flag, the part's own colour mapping (`c:clrMapOvr`, CT_ColorMapping),
// the print settings, the pivot chart format overrides and the user shapes reference. The print
// settings and pivot formats are kept as written (`printSettings`, `pivotFormatsXml`): an editor
// supplies the element it encoded, and the operation places it; `null` removes the element.
import { NS, buildXml, parseXml } from '../../xml/index';
import type { ChartSpace } from '../model';
import { editChartSpace } from './draft';

/** The theme colour roles a colour mapping assigns (CT_ColorMapping's required attributes). */
export const CHART_COLOR_MAP_ROLES = [
	'bg1',
	'tx1',
	'bg2',
	'tx2',
	'accent1',
	'accent2',
	'accent3',
	'accent4',
	'accent5',
	'accent6',
	'hlink',
	'folHlink',
] as const;

/** ST_ColorSchemeIndex: the theme colours a role may map to. */
const SCHEME_INDEXES = new Set([
	'dk1',
	'lt1',
	'dk2',
	'lt2',
	'accent1',
	'accent2',
	'accent3',
	'accent4',
	'accent5',
	'accent6',
	'hlink',
	'folHlink',
]);

/** Sets `c:externalData/c:autoUpdate`; a part without an external data reference is refused. */
export function setChartExternalDataAutoUpdate(space: ChartSpace, autoUpdate: boolean): ChartSpace {
	if (typeof autoUpdate !== 'boolean') throw new RangeError('Chart autoUpdate is a flag');
	if (space.externalDataRelId === undefined)
		throw new RangeError('The chart has no external data reference');
	return editChartSpace(space, (draft) => {
		draft.externalDataAutoUpdate = autoUpdate;
	});
}

/**
 * Sets the part's colour mapping (every role of CT_ColorMapping to a theme colour) or removes it
 * (`null`), so the chart follows the slide's mapping again. Child elements of an existing mapping
 * (its extension list) are kept.
 */
export function setChartColorMapOverride(
	space: ChartSpace,
	mapping: Readonly<Record<string, string>> | null,
): ChartSpace {
	if (mapping === null)
		return editChartSpace(space, (draft) => {
			delete draft.colorMapOverrideXml;
		});
	const roles = Object.keys(mapping);
	const known: readonly string[] = CHART_COLOR_MAP_ROLES;
	if (
		roles.length !== CHART_COLOR_MAP_ROLES.length ||
		roles.some((role) => !known.includes(role) || !SCHEME_INDEXES.has(mapping[role]!))
	)
		throw new RangeError('A chart colour mapping assigns each of the 12 roles a theme colour');
	const doc = parseXml(space.colorMapOverrideXml ?? `<c:clrMapOvr xmlns:c="${NS.c}"/>`, {
		label: 'Chart colour mapping',
	});
	const element = doc.documentElement;
	for (const attribute of [...element.attributes])
		if (!attribute.name.startsWith('xmlns')) element.removeAttribute(attribute.name);
	for (const role of CHART_COLOR_MAP_ROLES) element.setAttribute(role, mapping[role]!);
	const xml = buildXml(element);
	return editChartSpace(space, (draft) => {
		draft.colorMapOverrideXml = xml;
	});
}

/** Checks that `xml` is one `c:` element named `local`. */
function checkElement(xml: string, local: string): void {
	const root = parseXml(xml, { label: `Chart ${local}` }).documentElement;
	if (root.localName !== local || root.namespaceURI !== NS.c)
		throw new RangeError(`Expected a c:${local} element`);
}

/** Places encoded print settings (`c:printSettings`), or removes them (`null`). */
export function setChartPrintSettingsXml(space: ChartSpace, xml: string | null): ChartSpace {
	if (xml !== null) checkElement(xml, 'printSettings');
	return editChartSpace(space, (draft) => {
		if (xml === null) delete draft.printSettings;
		else draft.printSettings = xml;
	});
}

/** Places encoded pivot chart format overrides (`c:pivotFmts`), or removes them (`null`). */
export function setChartPivotFormatsXml(space: ChartSpace, xml: string | null): ChartSpace {
	if (xml !== null) checkElement(xml, 'pivotFmts');
	return editChartSpace(space, (draft) => {
		if (xml === null) delete draft.pivotFormatsXml;
		else draft.pivotFormatsXml = xml;
	});
}

/**
 * Points `c:userShapes` at the drawing part related as `relationshipId`, or removes the reference
 * (`null`). The package adapter writes the drawing part and its relationship.
 */
export function setChartUserShapesReference(
	space: ChartSpace,
	relationshipId: string | null,
): ChartSpace {
	if (relationshipId !== null && !/^[A-Za-z_][\w.-]*$/u.test(relationshipId))
		throw new RangeError('A relationship id is an XML name');
	return editChartSpace(space, (draft) => {
		if (relationshipId === null) delete draft.userShapesRelId;
		else draft.userShapesRelId = relationshipId;
	});
}
