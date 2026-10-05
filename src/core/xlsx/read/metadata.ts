// Cell metadata (`xl/metadata.xml`): which `cm` indices mark dynamic-array formulas.
import { elements, parseXml, type XmlElement } from '../../xml/index.js';
import { att } from './xml-util.js';

/** The `metadataType` name Excel uses for dynamic-array properties. */
export const DYNAMIC_ARRAY_TYPE = 'XLDAPR';

const childrenNamed = (parent: XmlElement | undefined, local: string): XmlElement[] =>
	parent ? elements(parent).filter((e) => e.localName === local) : [];

/**
 * The one-based `cm` values (cell metadata block indices) that mark a dynamic-array formula:
 * blocks whose record points at an `XLDAPR` future-metadata entry with `fDynamic="1"`.
 */
export function parseDynamicArrayMetadata(xml: string | undefined): Set<number> {
	const out = new Set<number>();
	if (!xml) return out;
	let root: XmlElement | null;
	try {
		root = parseXml(xml).documentElement;
	} catch {
		return out;
	}
	if (!root) return out;
	const top = new Map<string, XmlElement[]>();
	for (const child of elements(root))
		top.set(child.localName, [...(top.get(child.localName) ?? []), child]);
	const types = childrenNamed(top.get('metadataTypes')?.[0], 'metadataType').map(
		(t) => att(t, 'name') ?? '',
	);
	const future = (top.get('futureMetadata') ?? []).find(
		(f) => att(f, 'name') === DYNAMIC_ARRAY_TYPE,
	);
	const dynamic = childrenNamed(future, 'bk').map((bk) => {
		const props = bk.getElementsByTagNameNS('*', 'dynamicArrayProperties')[0];
		const flag = props?.getAttribute('fDynamic');
		return flag === '1' || flag === 'true';
	});
	childrenNamed(top.get('cellMetadata')?.[0], 'bk').forEach((bk, index) => {
		for (const rc of childrenNamed(bk, 'rc')) {
			const type = types[Number(att(rc, 't')) - 1];
			if (type === DYNAMIC_ARRAY_TYPE && dynamic[Number(att(rc, 'v'))]) out.add(index + 1);
		}
	});
	return out;
}
