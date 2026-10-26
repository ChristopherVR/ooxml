// Source-preserving colour edits. Existing transforms retain their attributes and namespaces.
import { NS, elements, parseXml, type XmlElement } from '../xml/index';
import { parseDrawingColor, parseDrawingColorIn } from './drawing-color';
import { drawingColorXml } from './write-color';
import type { DrawingColor } from './types';

export function patchDrawingColorIn(container: XmlElement, color: DrawingColor): void {
	const old = elements(container).find((node) => parseDrawingColor(node));
	const before = parseDrawingColorIn(container);
	if (!old || !before || before.kind !== color.kind) {
		const next = container.ownerDocument.importNode(
			parseXml(drawingColorXml(color)).documentElement,
			true,
		);
		if (old) container.replaceChild(next, old);
		else container.insertBefore(next, container.firstChild);
		return;
	}
	if (before.value !== color.value) {
		if (color.kind === 'scrgb' || color.kind === 'hsl') {
			const keys = color.kind === 'scrgb' ? ['r', 'g', 'b'] : ['hue', 'sat', 'lum'];
			const values = color.value.split(',');
			keys.forEach((key, index) => old.setAttribute(key, values[index] ?? '0'));
		} else old.setAttribute('val', color.value);
	}
	if (before.fallback !== color.fallback && color.kind === 'system') {
		if (color.fallback === undefined) old.removeAttribute('lastClr');
		else old.setAttribute('lastClr', color.fallback);
	}
	const originals = elements(old);
	color.transforms.forEach((transform, index) => {
		const source = originals[index];
		if (source?.localName === transform.name) {
			if (before.transforms[index]?.value !== transform.value) {
				if (transform.value === '') source.removeAttribute('val');
				else source.setAttribute('val', transform.value);
			}
			return;
		}
		if (!/^[A-Za-z_][\w.-]*$/u.test(transform.name))
			throw new Error('Invalid DrawingML colour transform');
		const next = old.ownerDocument.createElementNS(NS.a, `a:${transform.name}`);
		if (transform.value !== '') next.setAttribute('val', transform.value);
		if (source) old.replaceChild(next, source);
		else old.appendChild(next);
	});
	for (const extra of originals.slice(color.transforms.length)) old.removeChild(extra);
}
