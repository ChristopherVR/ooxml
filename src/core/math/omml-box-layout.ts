/** Lays out parsed OMML as boxes (see `omml-box.ts`). Unknown constructs lay out their children. */
import {
	emptyBox,
	glyphEm,
	isBinary,
	isRelation,
	row,
	round,
	scriptStyle,
	shifted,
	type Box,
	type MathStyle,
} from './omml-box';
import { layoutStructure } from './omml-box-structures';
import { child, ensureArray, readText, val } from './omml-mathml-helpers';
import type { OmmlNode } from './omml-node';
import { stripXmlOrderSuffix } from './omml-node';

const isOn = (value: string) => value === '1' || value === 'on' || value === 'true';

/** A run of text: letters are italic unless the run asks for upright, operators get spacing. */
function layoutRun(node: OmmlNode, style: MathStyle, atStart: boolean): Box {
	const text = readText(node);
	if (!text) return emptyBox();
	const rPr = child(node, 'm:rPr');
	const sty = val(child(rPr, 'm:sty'));
	const plain = style.plain || isOn(val(child(rPr, 'm:nor'))) || sty === 'p' || sty === 'b';
	const bold = Boolean(style.bold) || sty === 'b' || sty === 'bi';
	const boxes: Box[] = [];
	let index = 0;
	for (const ch of Array.from(text)) {
		const italic = !plain && /[A-Za-zα-ω]/.test(ch);
		const spaced = isRelation(ch) ? 0.28 : isBinary(ch) && !(atStart && index === 0) ? 0.22 : 0;
		const width = glyphEm(ch, bold) * style.size;
		const pad = spaced * style.size;
		const box = emptyBox();
		box.width = width + pad * 2;
		box.ascent = style.size * 0.78;
		box.descent = style.size * 0.22;
		if (ch !== ' ' || width)
			box.prims.push({
				kind: 'text',
				x: pad,
				y: 0,
				text: ch,
				size: style.size,
				italic,
				bold,
				fit: width,
			});
		boxes.push(box);
		index++;
	}
	return mergeText(row(boxes));
}

/** Joins neighbouring glyphs with the same style into one text primitive, fitted to its width. */
function mergeText(box: Box): Box {
	const merged: Box['prims'] = [];
	for (const prim of box.prims) {
		const last = merged.at(-1);
		if (
			prim.kind === 'text' &&
			last?.kind === 'text' &&
			last.italic === prim.italic &&
			last.bold === prim.bold &&
			last.size === prim.size &&
			last.y === prim.y &&
			Math.abs(last.x + last.fit - prim.x) < 0.01
		)
			merged[merged.length - 1] = {
				...last,
				text: last.text + prim.text,
				fit: last.fit + prim.fit,
			};
		else merged.push(prim);
	}
	return { ...box, prims: merged };
}

/** The boxes of every content child of `node`, in document order. */
export function layoutChildren(node: OmmlNode, style: MathStyle): Box {
	const boxes: Box[] = [];
	for (const key of Object.keys(node)) {
		if (key.startsWith('@_') || key === '#text') continue;
		const tag = stripXmlOrderSuffix(key);
		for (const item of ensureArray(node[key]))
			boxes.push(layoutElement(tag, item, style, boxes.length === 0));
	}
	return row(boxes);
}

function layoutElement(tag: string, node: OmmlNode, style: MathStyle, atStart: boolean): Box {
	if (tag === 'm:r') return layoutRun(node, style, atStart);
	if (tag === 'm:oMath' || tag === 'm:oMathPara') return layoutChildren(node, style);
	if (/Pr$/.test(tag) || tag === 'm:ctrlPr') return emptyBox();
	return layoutStructure(tag, node, style, { layout: layoutChildren, script: scriptStyle });
}

/** Result of drawing an equation: size in px and the descent below the text baseline. */
export interface EquationMetrics {
	box: Box;
	widthPx: number;
	heightPx: number;
	descentPx: number;
}

/** Lays out an `m:oMath` / `m:oMathPara` tree at `sizePx`. */
export function layoutEquation(root: OmmlNode, sizePx: number): EquationMetrics {
	const box = layoutChildren(root, { level: 0, size: sizePx });
	const pad = 1;
	const ascent = Math.max(box.ascent, sizePx * 0.78) + pad;
	const descent = Math.max(box.descent, sizePx * 0.22) + pad;
	return {
		box: { ...box, prims: shifted(box.prims, pad, 0), ascent, descent },
		widthPx: round(Math.max(box.width, sizePx * 0.5) + pad * 2),
		heightPx: round(ascent + descent),
		descentPx: round(descent),
	};
}
