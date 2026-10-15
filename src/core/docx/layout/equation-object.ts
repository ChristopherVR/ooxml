// Print Layout equations: Office Math drawn as SVG by the shared math box model, so the size the
// engine reserves on the line is the size that is painted.
import { convertOmmlToLatex, drawEquation, parseOmml } from '../../math/index';
import type { LayoutObject } from './input';

const cache = new Map<string, LayoutObject | null>();
const CACHE_LIMIT = 256;

function latexLabel(omml: string): string {
	try {
		return convertOmmlToLatex(parseOmml(omml));
	} catch {
		return 'Equation';
	}
}

/**
 * The inline object for an equation at `sizePx`, or undefined when its markup holds nothing that
 * can be drawn (the caller then shows a marker).
 */
export function equationObject(omml: string, sizePx: number): LayoutObject | undefined {
	const key = `${sizePx}|${omml}`;
	const hit = cache.get(key);
	if (hit !== undefined) return hit ?? undefined;
	const drawing = drawEquation(omml, sizePx, latexLabel(omml));
	const object: LayoutObject | null = drawing
		? {
				partName: '',
				contentType: '',
				widthPx: drawing.widthPx,
				heightPx: drawing.heightPx,
				descentPx: drawing.descentPx,
				chart: { svg: drawing.svg, label: 'Equation' },
			}
		: null;
	if (cache.size >= CACHE_LIMIT) cache.delete(cache.keys().next().value as string);
	cache.set(key, object);
	return object ?? undefined;
}
