/**
 * The relationship ids a slide element names, read and renamed in one place.
 *
 * An element names the parts it shows through ids that only mean something on the part that
 * declares them: `r:id` / `r:embed` / `r:link` / `r:dm` ... attributes in its markup, plus the
 * copies of those ids the typed model keeps (a chart's `chartRelationshipId`, a diagram's
 * `dataRelId` and siblings, the last also naming the drawing relationship that only the data
 * part cites). Markup of other parts (a chart's own XML, an ink part) is not read: its ids belong
 * to that part.
 *
 * @module core/core/runtime/element-relationship-ids
 */
import type { PptxElement, XmlObject } from '../../types';
import { isTemplateElementId } from '../../utils/group-ops';
import { GROUP_CHILD_TAGS } from './save-group-shape-xml';

/** Element fields holding markup of the element's own part. */
const OWN_MARKUP_KEYS = ['rawXml', 'rawMediaReferenceXml', 'extLstXml'] as const;

/** A group's children are elements of their own; its markup also holds them. */
const GROUP_CHILD_KEYS: ReadonlySet<string> = new Set([...GROUP_CHILD_TAGS, 'mc:AlternateContent']);

const SMART_ART_ID_KEYS = [
	'dataRelId',
	'layoutRelId',
	'styleRelId',
	'colorsRelId',
	'drawingRelId',
] as const;

const isRelationshipAttribute = (key: string): boolean => key.startsWith('@_r:');

function visitMarkup(
	node: unknown,
	skipKeys: ReadonlySet<string> | undefined,
	visit: (holder: XmlObject, key: string) => void,
): void {
	if (Array.isArray(node)) {
		for (const item of node) visitMarkup(item, skipKeys, visit);
		return;
	}
	if (!node || typeof node !== 'object') return;
	const record = node as XmlObject;
	for (const key of Object.keys(record)) {
		if (skipKeys?.has(key)) continue;
		const value = record[key];
		if (isRelationshipAttribute(key) && typeof value === 'string') visit(record, key);
		else if (value && typeof value === 'object') visitMarkup(value, undefined, visit);
	}
}

function visitOwnMarkup(
	element: PptxElement,
	visit: (holder: XmlObject, key: string) => void,
): void {
	const skip = element.type === 'group' ? GROUP_CHILD_KEYS : undefined;
	const fields = element as unknown as Record<string, unknown>;
	for (const key of OWN_MARKUP_KEYS) visitMarkup(fields[key], skip, visit);
}

/** Typed id holders: the object and key of each id copy the model keeps outside markup. */
function typedIdHolders(element: PptxElement): Array<[Record<string, unknown>, string]> {
	if (element.type === 'chart' && element.chartData)
		return [[element.chartData as unknown as Record<string, unknown>, 'chartRelationshipId']];
	if (element.type === 'smartArt' && element.smartArtData) {
		const data = element.smartArtData as unknown as Record<string, unknown>;
		return SMART_ART_ID_KEYS.map((key) => [data, key]);
	}
	return [];
}

/** Every relationship id the element names on its own part. */
export function elementRelationshipIds(element: PptxElement): Set<string> {
	const ids = new Set<string>();
	visitOwnMarkup(element, (holder, key) => {
		const value = holder[key] as string;
		if (value) ids.add(value);
	});
	for (const [holder, key] of typedIdHolders(element)) {
		const value = holder[key];
		if (typeof value === 'string' && value) ids.add(value);
	}
	return ids;
}

/**
 * Rename relationship ids on an element in place. The markup trees are copied first: clones share
 * them by reference (`clone-utils`), and the source element must keep its own ids.
 */
export function renameElementRelationshipIds(
	element: PptxElement,
	renames: ReadonlyMap<string, string>,
): void {
	if (renames.size === 0) return;
	const fields = element as unknown as Record<string, unknown>;
	for (const key of OWN_MARKUP_KEYS)
		if (fields[key] !== undefined) fields[key] = structuredClone(fields[key]);
	visitOwnMarkup(element, (holder, key) => {
		const renamed = renames.get(holder[key] as string);
		if (renamed) holder[key] = renamed;
	});
	for (const [holder, key] of typedIdHolders(element)) {
		const renamed = renames.get(holder[key] as string);
		if (renamed) holder[key] = renamed;
	}
}

/**
 * Record `part` as the relationship source of every element (group members too) that names ids.
 * Layout and master elements shown on a slide belong to their own part and are left alone.
 */
export function recordRelationshipSource(elements: readonly PptxElement[], part: string): void {
	for (const element of elements) {
		if (isTemplateElementId(element.id)) continue;
		if (element.type === 'group') recordRelationshipSource(element.children, part);
		if (elementRelationshipIds(element).size > 0) element.relationshipSourcePath = part;
	}
}
