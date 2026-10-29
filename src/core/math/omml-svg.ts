/** Draws Office Math as SVG sized by the box model: the same numbers lay out and paint it. */
import { round, type Prim } from './omml-box';
import { layoutEquation } from './omml-box-layout';
import { escapeXml } from './omml-mathml-helpers';
import { parseOmml } from './omml-dom';
import type { OmmlNode } from './omml-node';

const FONT_FAMILY = "'Cambria Math','STIX Two Math','Latin Modern Math','Times New Roman',serif";

export interface EquationDrawing {
	/** Standalone `<svg>` markup painted in `currentColor`. */
	svg: string;
	widthPx: number;
	heightPx: number;
	/** How far the equation extends below the text baseline, for baseline alignment. */
	descentPx: number;
}

function primMarkup(prim: Prim): string {
	switch (prim.kind) {
		case 'text': {
			const attrs = [
				`font-size="${round(prim.size)}"`,
				prim.italic ? 'font-style="italic"' : '',
				prim.bold ? 'font-weight="bold"' : '',
				Array.from(prim.text).length > 1
					? `textLength="${round(prim.fit)}" lengthAdjust="spacing"`
					: '',
			].filter(Boolean);
			const at = prim.stretch
				? `transform="translate(${round(prim.x)} ${round(prim.y)}) scale(1 ${round(prim.stretch)})"`
				: `x="${round(prim.x)}" y="${round(prim.y)}"`;
			return `<text ${at} ${attrs.join(' ')}>${escapeXml(prim.text)}</text>`;
		}
		case 'line':
			return `<line x1="${round(prim.x1)}" y1="${round(prim.y1)}" x2="${round(prim.x2)}" y2="${round(prim.y2)}" stroke="currentColor" stroke-width="${round(prim.width)}"/>`;
		case 'rect':
			return `<rect x="${round(prim.x)}" y="${round(prim.y)}" width="${round(prim.w)}" height="${round(prim.h)}" fill="none" stroke="currentColor" stroke-width="${round(prim.width)}"/>`;
		default:
			return `<path d="${prim.d}" fill="none" stroke="currentColor" stroke-width="${round(prim.width)}" stroke-linejoin="round"/>`;
	}
}

/** Draws an already parsed OMML tree at `sizePx`. Returns undefined when it holds no content. */
export function drawOmmlTree(
	root: OmmlNode,
	sizePx: number,
	label = '',
	options: { display?: boolean } = {},
): EquationDrawing | undefined {
	const { box, widthPx, heightPx, descentPx } = layoutEquation(root, sizePx, options);
	if (!box.prims.length) return undefined;
	const body = box.prims.map(primMarkup).join('');
	const aria = label ? ` aria-label="${escapeXml(label).replace(/"/g, '&quot;')}"` : '';
	const svg =
		`<svg xmlns="http://www.w3.org/2000/svg" width="${widthPx}" height="${heightPx}" viewBox="0 0 ${widthPx} ${heightPx}" role="img"${aria}>` +
		`<g font-family="${FONT_FAMILY}" fill="currentColor" transform="translate(0 ${round(box.ascent)})">${body}</g></svg>`;
	return { svg, widthPx, heightPx, descentPx };
}

/** Draws OMML markup (`m:oMath` or `m:oMathPara`); undefined when it cannot be read or is empty. */
export function drawEquation(
	omml: string,
	sizePx: number,
	label = '',
	options: { display?: boolean } = {},
): EquationDrawing | undefined {
	try {
		return drawOmmlTree(parseOmml(omml), sizePx, label, options);
	} catch {
		return undefined;
	}
}
