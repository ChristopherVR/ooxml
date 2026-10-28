import { children, first, NS, parseXml, type XmlElement } from '../../xml';
import type { RegionValueColorPosition } from './types';
function schemaPosition(
	node: XmlElement | undefined,
	end: 'min' | 'mid' | 'max',
): RegionValueColorPosition {
	for (const kind of ['number', 'percent'] as const) {
		const element = first(node, kind, NS.cx),
			raw = element?.getAttribute('val');
		const value = raw === null || raw === undefined || raw === '' ? NaN : Number(raw);
		if (Number.isFinite(value)) return { kind, value };
	}
	return end === 'mid' ? { kind: 'percent', value: 50 } : { kind: end };
}
/** CT_ValueColorPositions min/mid/max children, retaining the earlier cx:pos shorthand. */
export function parseRegionValueColorPositions(
	xml: string | undefined,
): RegionValueColorPosition[] | undefined {
	if (!xml) return undefined;
	const root = parseXml(`<cx:root xmlns:cx="${NS.cx}">${xml}</cx:root>`).documentElement;
	const positions = first(root, 'valueColorPositions', NS.cx);
	if (!positions) return undefined;
	const count = positions.getAttribute('count');
	const native = ['min', 'mid', 'max'].some((key) => first(positions, key, NS.cx));
	if (native || (count !== null && children(positions, 'pos', NS.cx).length === 0)) {
		const keys =
			count === '3' || first(positions, 'mid', NS.cx)
				? (['min', 'mid', 'max'] as const)
				: (['min', 'max'] as const);
		return keys.map((key) => schemaPosition(first(positions, key, NS.cx), key));
	}
	const legacy = children(positions, 'pos', NS.cx).flatMap(
		(element): RegionValueColorPosition[] => {
			const type = element.getAttribute('type');
			if (type === 'min' || type === 'max') return [{ kind: type }];
			if (type !== 'num' && type !== 'percent') return [];
			const raw = element.getAttribute('val'),
				value = raw === null || raw === '' ? NaN : Number(raw);
			return [
				{ kind: type === 'num' ? 'number' : type, ...(Number.isFinite(value) ? { value } : {}) },
			];
		},
	);
	return legacy.length ? legacy : undefined;
}
