import { resolveDrawingOuterShadow } from 'ooxml-core/drawingml';
import type { DiagramDrawingShape } from 'ooxml-core/diagram';
import type { SchemeColors, SmartArtRenderReport } from './smartart-svg';

const NS = 'http://www.w3.org/2000/svg';
let serial = 0;

export function diagramEffects(
	doc: Document,
	defs: SVGDefsElement,
	body: SVGElement,
	shape: DiagramDrawingShape,
	scheme: SchemeColors,
	report: SmartArtRenderReport,
): void {
	for (const effect of shape.effects?.effects ?? []) {
		const shadow =
			effect.kind === 'outerShadow'
				? resolveDrawingOuterShadow(effect, { scheme: (name) => scheme[name] })
				: undefined;
		if (!shadow) {
			report.unappliedEffects = [...new Set([...(report.unappliedEffects ?? []), effect.kind])];
			continue;
		}
		const filter = doc.createElementNS(NS, 'filter');
		filter.id = `office-smartart-shadow-${++serial}`;
		filter.setAttribute('filterUnits', 'userSpaceOnUse');
		const pad = Math.max(
			shadow.blur * 3 + Math.abs(shadow.dx),
			shadow.blur * 3 + Math.abs(shadow.dy),
			1,
		);
		filter.setAttribute('x', String(-pad));
		filter.setAttribute('y', String(-pad));
		filter.setAttribute('width', String(shape.frame.width / 9525 + pad * 2));
		filter.setAttribute('height', String(shape.frame.height / 9525 + pad * 2));
		filter.setAttribute('color-interpolation-filters', 'sRGB');
		const drop = doc.createElementNS(NS, 'feDropShadow');
		drop.setAttribute('dx', String(shadow.dx));
		drop.setAttribute('dy', String(shadow.dy));
		drop.setAttribute('stdDeviation', String(shadow.blur / 2));
		drop.setAttribute('flood-color', shadow.color);
		drop.setAttribute('flood-opacity', String(shadow.opacity));
		filter.append(drop);
		defs.append(filter);
		body.setAttribute('filter', `url(#${filter.id})`);
	}
}
