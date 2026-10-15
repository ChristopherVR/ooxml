/** Structured Office Math constructs as boxes: fractions, scripts, radicals, n-ary, delimiters... */
import {
	AXIS,
	emptyBox,
	glyph,
	isOn,
	overlay,
	placed,
	round,
	row,
	scripts,
	shifted,
	type Box,
	type LayoutContext,
	type MathStyle,
} from './omml-box';
import { above, below, delimiter, matrix, nary } from './omml-box-grouped';
import { child, ensureArray, val } from './omml-mathml-helpers';
import type { OmmlNode } from './omml-node';

export type { LayoutContext } from './omml-box';

function fraction(node: OmmlNode, style: MathStyle, ctx: LayoutContext): Box {
	const type = val(child(child(node, 'm:fPr'), 'm:type'));
	const inner = ctx.script({ ...style, level: Math.min(style.level, 0) });
	const small = style.level === 0 ? { ...style, size: style.size * 0.9 } : inner;
	const num = ctx.layout(child(node, 'm:num'), small);
	const den = ctx.layout(child(node, 'm:den'), small);
	if (type === 'lin') {
		const slash = glyph('/', style.size);
		return row([num, slash, den]);
	}
	const gap = style.size * 0.12;
	const axis = style.size * AXIS;
	const width = Math.max(num.width, den.width) + style.size * 0.2;
	const bar =
		type === 'noBar'
			? []
			: [
					{
						kind: 'line' as const,
						x1: 0,
						x2: width,
						y1: -axis,
						y2: -axis,
						width: Math.max(0.8, style.size * 0.05),
					},
				];
	const box = overlay([
		{ box: num, x: (width - num.width) / 2, y: -(axis + gap + num.descent) },
		{ box: den, x: (width - den.width) / 2, y: -axis + gap + den.ascent },
	]);
	box.width = width;
	box.prims.push(...bar);
	return box;
}

function radical(node: OmmlNode, style: MathStyle, ctx: LayoutContext): Box {
	const content = ctx.layout(child(node, 'm:e'), style);
	const hide = isOn(val(child(child(node, 'm:radPr'), 'm:degHide')));
	const degree = hide
		? undefined
		: ctx.layout(child(node, 'm:deg'), { ...ctx.script(ctx.script(style)) });
	const gap = style.size * 0.12;
	const top = -(content.ascent + gap);
	const bottom = content.descent + style.size * 0.05;
	const height = bottom - top;
	const sign = style.size * 0.55;
	const lead = degree && degree.width ? Math.max(0, degree.width - sign * 0.4) : 0;
	const x0 = lead;
	const d = `M${round(x0)} ${round(top + height * 0.58)} L${round(x0 + sign * 0.22)} ${round(top + height * 0.52)} L${round(x0 + sign * 0.5)} ${round(bottom)} L${round(x0 + sign)} ${round(top)} L${round(x0 + sign + content.width + style.size * 0.1)} ${round(top)}`;
	const parts = [{ box: content, x: x0 + sign + style.size * 0.05, y: 0 }];
	const box = overlay(parts);
	box.ascent = Math.max(box.ascent, -top);
	box.descent = Math.max(box.descent, bottom);
	box.width = x0 + sign + content.width + style.size * 0.15;
	box.prims.push({ kind: 'path', d, width: Math.max(0.8, style.size * 0.06) });
	if (degree) box.prims.push(...shifted(degree.prims, 0, top + height * 0.35 - degree.descent));
	return box;
}

/** Lays out one structured element; unknown tags lay out their content children. */
export function layoutStructure(
	tag: string,
	node: OmmlNode,
	style: MathStyle,
	ctx: LayoutContext,
): Box {
	const arg = (name: string, next = style) => ctx.layout(child(node, name), next);
	const small = () => ctx.script(style);
	switch (tag) {
		case 'm:f':
			return fraction(node, style, ctx);
		case 'm:rad':
			return radical(node, style, ctx);
		case 'm:sSup':
			return scripts(arg('m:e'), undefined, arg('m:sup', small()), style);
		case 'm:sSub':
			return scripts(arg('m:e'), arg('m:sub', small()), undefined, style);
		case 'm:sSubSup':
			return scripts(arg('m:e'), arg('m:sub', small()), arg('m:sup', small()), style);
		case 'm:sPre': {
			const pre = scripts(emptyBox(), arg('m:sub', small()), arg('m:sup', small()), style);
			return row([pre, arg('m:e')], 1);
		}
		case 'm:nary':
			return nary(node, style, ctx);
		case 'm:d':
			return delimiter(node, style, ctx);
		case 'm:m':
			return matrix(node, style, ctx);
		case 'm:acc': {
			const chr = val(child(child(node, 'm:accPr'), 'm:chr')) || '̂';
			const mark = glyph(
				chr === '̂' ? '^' : chr === '̃' ? '~' : chr === '̄' ? '¯' : chr,
				style.size * 0.8,
			);
			mark.ascent = style.size * 0.2;
			mark.descent = 0;
			return above(arg('m:e'), mark, style);
		}
		case 'm:bar': {
			const base = arg('m:e');
			const under = val(child(child(node, 'm:barPr'), 'm:pos')) === 'bot';
			const y = under ? base.descent + style.size * 0.08 : -(base.ascent + style.size * 0.08);
			const box = overlay([{ box: base, x: 0, y: 0 }]);
			box.prims.push({
				kind: 'line',
				x1: 0,
				x2: base.width,
				y1: y,
				y2: y,
				width: Math.max(0.8, style.size * 0.05),
			});
			box.ascent = Math.max(box.ascent, under ? 0 : -y);
			box.descent = Math.max(box.descent, under ? y : 0);
			return box;
		}
		case 'm:limLow':
			return below(arg('m:e'), arg('m:lim', small()), style);
		case 'm:limUpp':
			return above(arg('m:e'), arg('m:lim', small()), style);
		case 'm:groupChr': {
			const pr = child(node, 'm:groupChrPr');
			const mark = glyph(val(child(pr, 'm:chr')) || '⏟', style.size);
			mark.ascent = style.size * 0.2;
			mark.descent = style.size * 0.1;
			return val(child(pr, 'm:pos')) === 'top'
				? above(arg('m:e'), mark, style)
				: below(arg('m:e'), mark, style);
		}
		case 'm:eqArr': {
			const lines = ensureArray(node['m:e']).map((line) => ctx.layout(line, style));
			const result = emptyBox();
			const gap = style.size * 0.25;
			const total =
				lines.reduce((sum, line) => sum + line.ascent + line.descent, 0) +
				gap * Math.max(0, lines.length - 1);
			let y = -(total / 2) + style.size * AXIS;
			for (const line of lines) {
				y += line.ascent;
				result.prims.push(...shifted(line.prims, 0, y));
				result.width = Math.max(result.width, line.width);
				y += line.descent + gap;
			}
			result.ascent = total / 2 + style.size * AXIS;
			result.descent = total / 2 - style.size * AXIS;
			return result;
		}
		case 'm:func': {
			const name = arg('m:fName', { ...style, plain: true });
			return row([name, arg('m:e')], style.size * 0.17);
		}
		case 'm:borderBox': {
			const inner = arg('m:e');
			const pad = style.size * 0.15;
			const box = placed(inner, pad, 0);
			box.width = inner.width + pad * 2;
			box.ascent += pad;
			box.descent += pad;
			box.prims.push({
				kind: 'rect',
				x: 0,
				y: -box.ascent,
				w: box.width,
				h: box.ascent + box.descent,
				width: 0.8,
			});
			return box;
		}
		case 'm:box':
		case 'm:phant':
			return arg('m:e');
		default:
			return ctx.layout(node, style);
	}
}
