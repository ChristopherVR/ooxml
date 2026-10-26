import { NS, buildXml, first, parseXml } from '../xml';
import { cBool, type ChartParseContext } from './parse-util';
import { CHART_ROOT_BINDINGS } from './xml-fragment';
import type { XmlElement } from '../xml';
import { raw, type ChartWriteContext } from './write-util';

export const CHART_PROTECTION_FIELDS = [
	'chartObject',
	'data',
	'formatting',
	'selection',
	'userInterface',
] as const;
export type ChartProtection = Partial<Record<(typeof CHART_PROTECTION_FIELDS)[number], boolean>> & {
	sourceXml?: string;
};

export function parseChartProtection(
	context: ChartParseContext,
	node: XmlElement | undefined,
): ChartProtection | undefined {
	if (!node) return undefined;
	const protection: ChartProtection = { sourceXml: buildXml(node) };
	for (const key of CHART_PROTECTION_FIELDS) {
		const value = cBool(context, node, key);
		if (value !== undefined) protection[key] = value;
	}
	return protection;
}

export function chartProtectionXml(
	context: ChartWriteContext,
	protection: ChartProtection | null | undefined,
): string {
	if (!protection) return '';
	const declarations = [...CHART_ROOT_BINDINGS].map(([p, uri]) => ` xmlns:${p}="${uri}"`).join('');
	const root = parseXml(protection.sourceXml ?? `<c:protection${declarations}/>`).documentElement;
	for (const key of CHART_PROTECTION_FIELDS) {
		const old = first(root, key, NS.c);
		const value = protection[key];
		if (value === undefined) {
			if (old) root.removeChild(old);
			continue;
		}
		const previous = old?.getAttribute('val');
		if (old && (previous === '1' || previous === 'true' || previous === null) === value) continue;
		const node = old ?? root.ownerDocument.createElementNS(NS.c, `c:${key}`);
		node.setAttribute('val', value ? '1' : '0');
		if (!old) {
			const later = CHART_PROTECTION_FIELDS.slice(CHART_PROTECTION_FIELDS.indexOf(key) + 1)
				.map((name) => first(root, name, NS.c))
				.find(Boolean);
			root.insertBefore(node, later ?? null);
		}
	}
	return raw(context, buildXml(root));
}
