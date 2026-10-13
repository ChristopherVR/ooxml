import {
	NS,
	buildXml,
	children,
	elements,
	first,
	parseXml,
	type XmlElement,
} from '../../xml/index';
import type { SparklineGroup, Worksheet } from '../model';
import { readSparklineGroup } from '../read/sparklines';
import { colorXml } from '../write/style-xml';

export const XM = 'http://schemas.microsoft.com/office/excel/2006/main';
// Excel treats extension URI strings case-sensitively, even though this one looks like a GUID.
const URI = '{05C60535-1F16-4fd2-B633-F4F36F0B64E0}';
export type SparklineGroupPatch = Partial<
	Omit<SparklineGroup, 'sparklines' | 'dateAxis' | 'dateFormula'>
>;
const COLORS = [
	'colorSeries',
	'colorNegative',
	'colorAxis',
	'colorMarkers',
	'colorFirst',
	'colorLast',
	'colorHigh',
	'colorLow',
] as const;
const SCALARS = [
	'type',
	'lineWeight',
	'markers',
	'high',
	'low',
	'first',
	'last',
	'negative',
	'displayEmptyCellsAs',
	'displayXAxis',
	'displayHidden',
	'rightToLeft',
	'minAxisType',
	'maxAxisType',
	'manualMin',
	'manualMax',
] as const;

/** JS callers can supply extra or undefined properties; only the supported fields enter XML. */
export function cleanSparklinePatch(patch: SparklineGroupPatch): SparklineGroupPatch {
	return Object.fromEntries(
		[...SCALARS, ...COLORS].flatMap((key) => (patch[key] === undefined ? [] : [[key, patch[key]]])),
	) as SparklineGroupPatch;
}

/** Patches only supported fields, retaining other attributes and extension content. */
export function patchSparklineNode(node: XmlElement, patch: SparklineGroupPatch): void {
	for (const key of SCALARS) {
		const value = patch[key];
		if (value === undefined) continue;
		node.setAttribute(key, typeof value === 'boolean' ? (value ? '1' : '0') : String(value));
	}
	// A change back from a custom scale removes its schema-incompatible manual bound.
	if (patch.minAxisType && patch.minAxisType !== 'custom') node.removeAttribute('manualMin');
	if (patch.maxAxisType && patch.maxAxisType !== 'custom') node.removeAttribute('manualMax');
	for (const key of COLORS) {
		const color = patch[key];
		if (!color) continue;
		const fresh = parseXml(
			`<x14:root xmlns:x14="${NS.x14}">${colorXml(color, `x14:${key}`)}</x14:root>`,
		).documentElement.firstChild!;
		const replacement = node.ownerDocument!.importNode(fresh, true);
		const old = first(node, key, NS.x14);
		if (old) node.replaceChild(replacement, old);
		else {
			const before = elements(node).find((child) => {
				const index = COLORS.indexOf(child.localName as (typeof COLORS)[number]);
				return (
					index > COLORS.indexOf(key) || child.localName === 'f' || child.localName === 'sparklines'
				);
			});
			node.insertBefore(replacement, before ?? null);
		}
	}
}

/** Rewrites only fragments whose sparkline groups changed. */
export function editSparklineXml(
	sheet: Worksheet,
	edit: (node: XmlElement, group: SparklineGroup) => boolean,
): string[] {
	return (sheet.preserved.get('extLst') ?? []).map((xml) => {
		if (!xml.includes('sparklineGroup')) return xml;
		const doc = parseXml(xml);
		let changed = false;
		for (const node of Array.from(doc.getElementsByTagNameNS(NS.x14, 'sparklineGroup'))) {
			if (!edit(node, readSparklineGroup(node))) continue;
			changed = true;
			const lines = first(node, 'sparklines', NS.x14);
			if (!lines || !children(lines, 'sparkline', NS.x14).length) {
				const groups = node.parentNode!;
				groups.removeChild(node);
				if (!children(groups, 'sparklineGroup', NS.x14).length)
					groups.parentNode!.removeChild(groups);
			}
		}
		return changed ? buildXml(doc.documentElement) : xml;
	});
}

/** Adds a group to the existing extension container rather than creating a second extLst. */
export function appendSparklineXml(sheet: Worksheet, groupXml: string): string[] {
	const source = [...(sheet.preserved.get('extLst') ?? [])];
	const found = source.findIndex((xml) => xml.includes('sparklineGroups'));
	const index = found >= 0 ? found : 0;
	const doc = parseXml(source[index] ?? `<extLst xmlns="${NS.x}"/>`);
	const root = doc.documentElement;
	let ext = children(root, 'ext', NS.x).find(
		(node) => node.getAttribute('uri')?.toUpperCase() === URI.toUpperCase(),
	);
	if (!ext) {
		ext = doc.createElementNS(NS.x, 'ext');
		ext.setAttribute('uri', URI);
		ext.setAttribute('xmlns:x14', NS.x14);
		root.appendChild(ext);
	}
	// Match Excel's known URI when reusing a container with a different GUID spelling.
	ext.setAttribute('uri', URI);
	let groups = first(ext, 'sparklineGroups', NS.x14);
	if (!groups) {
		groups = doc.createElementNS(NS.x14, 'x14:sparklineGroups');
		groups.setAttribute('xmlns:xm', XM);
		ext.appendChild(groups);
	}
	const fresh = doc.importNode(parseXml(groupXml).documentElement, true) as XmlElement;
	const revision = 'http://schemas.microsoft.com/office/spreadsheetml/2015/revision2';
	if (
		Array.from(doc.getElementsByTagNameNS(NS.x14, 'sparklineGroup')).some((node) =>
			node.hasAttributeNS(revision, 'uid'),
		)
	) {
		fresh.setAttributeNS('http://www.w3.org/2000/xmlns/', 'xmlns:xr2', revision);
		fresh.setAttributeNS(revision, 'xr2:uid', `{${globalThis.crypto.randomUUID().toUpperCase()}}`);
	}
	groups.appendChild(fresh);
	source[index] = buildXml(root);
	return source;
}
