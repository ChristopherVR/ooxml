import { NS, buildXml, parseXml } from '../xml/index';
import type { DiagramColor } from './types';

const ELEMENTS = {
	srgb: 'srgbClr',
	scheme: 'schemeClr',
	system: 'sysClr',
	preset: 'prstClr',
	scrgb: 'scrgbClr',
	hsl: 'hslClr',
} as const;

/** One DrawingML color choice, with its ordered transforms and system-color fallback. */
export function drawingColorXml(color: DiagramColor): string {
	const local = ELEMENTS[color.kind];
	const doc = parseXml(`<a:${local} xmlns:a="${NS.a}"/>`);
	const root = doc.documentElement;
	if (color.kind === 'scrgb' || color.kind === 'hsl') {
		const keys = color.kind === 'scrgb' ? ['r', 'g', 'b'] : ['hue', 'sat', 'lum'];
		const values = color.value.split(',');
		keys.forEach((key, i) => root.setAttribute(key, values[i] ?? '0'));
	} else root.setAttribute('val', color.value);
	if (color.fallback !== undefined && color.kind === 'system')
		root.setAttribute('lastClr', color.fallback);
	for (const transform of color.transforms) {
		if (!/^[A-Za-z_][\w.-]*$/.test(transform.name))
			throw new Error(`Invalid DrawingML transform ${transform.name}`);
		const node = doc.createElementNS(NS.a, `a:${transform.name}`);
		if (transform.value !== '') node.setAttribute('val', transform.value);
		root.appendChild(node);
	}
	return buildXml(doc);
}
