// Draws a worksheet shape, connector or group as SVG from the shared preset geometry tables
// (`geometry/preset-shape-evaluator`), the shared custom-geometry path data and the DrawingML
// fill, outline and text body. The Excel UI only paints this string (docs/xlsx-parity-status.md).
import { drawingPathData } from '../../drawingml/path-data';
import type { DrawingTextBody, DrawingTextParagraph } from '../../drawingml/types';
import { evaluatePresetShape, lookupPresetShape } from '../../geometry/preset-shape-evaluator';
import type { ThemePalette } from '../model';
import type { ShapeNode } from '../model-shapes';
import { EMU_PER_PX, shapeColor, shapeFill, shapeStroke } from './shape-paint';

interface Box {
	x: number;
	y: number;
	w: number;
	h: number;
}

interface Context {
	theme: ThemePalette | undefined;
	/** Pixels per pixel at 100%: the sheet zoom / 100. */
	scale: number;
	prefix: string;
	counter: { n: number };
	defs: string[];
}

const escapeXml = (text: string) =>
	text
		.replace(/&/g, '&amp;')
		.replace(/</g, '&lt;')
		.replace(/>/g, '&gt;')
		.replace(/"/g, '&quot;')
		.replace(/'/g, '&#39;');
const round = (value: number) => Math.round(value * 100) / 100;
const POINT_PX = 96 / 72;
const DEFAULT_FONT_PT = 11;

/** Greedy wrap with an average glyph width estimate (the core has no text measurer). */
export function wrapShapeLine(text: string, widthPx: number, fontPx: number): string[] {
	const perLine = Math.max(1, Math.floor(widthPx / (fontPx * 0.52)));
	if (text.length <= perLine) return [text];
	const lines: string[] = [];
	let current = '';
	for (const word of text.split(/\s+/)) {
		if (current && current.length + 1 + word.length > perLine) {
			lines.push(current);
			current = word;
		} else current = current ? `${current} ${word}` : word;
	}
	lines.push(current);
	return lines;
}

function paragraphText(body: DrawingTextBody, paragraph: DrawingTextParagraph) {
	const first = paragraph.runs[0];
	return {
		text: paragraph.runs.map((run) => run.text).join(''),
		sizePt: first?.sizePt ?? paragraph.defaultProperties?.sizePt ?? DEFAULT_FONT_PT,
		bold: first?.bold ?? paragraph.defaultProperties?.bold ?? false,
		italic: first?.italic ?? paragraph.defaultProperties?.italic ?? false,
		color: first?.color ?? paragraph.defaultProperties?.color,
		align: paragraph.align ?? 'l',
		body,
	};
}

function textSvg(node: ShapeNode, rect: Box, ctx: Context): string {
	const body = node.text;
	if (!body) return '';
	const left = ((body.insetsEmu?.left ?? 91440) / EMU_PER_PX) * ctx.scale;
	const right = ((body.insetsEmu?.right ?? 91440) / EMU_PER_PX) * ctx.scale;
	const top = ((body.insetsEmu?.top ?? 45720) / EMU_PER_PX) * ctx.scale;
	const bottom = ((body.insetsEmu?.bottom ?? 45720) / EMU_PER_PX) * ctx.scale;
	const innerW = Math.max(1, rect.w - left - right);
	const lines: {
		text: string;
		fontPx: number;
		color: string;
		bold: boolean;
		italic: boolean;
		align: string;
	}[] = [];
	for (const paragraph of body.paragraphs) {
		const info = paragraphText(body, paragraph);
		const fontPx = info.sizePt * POINT_PX * ctx.scale;
		const color =
			shapeColor(info.color, ctx.theme) ??
			shapeColor(node.style?.font?.color, ctx.theme) ??
			'#000000';
		for (const text of wrapShapeLine(info.text, innerW, fontPx))
			lines.push({ text, fontPx, color, bold: info.bold, italic: info.italic, align: info.align });
	}
	if (!lines.length) return '';
	const heights = lines.map((line) => line.fontPx * 1.2);
	const total = heights.reduce((sum, value) => sum + value, 0);
	const anchor = body.anchor ?? 't';
	let y =
		rect.y +
		(anchor === 'ctr' ? (rect.h - total) / 2 : anchor === 'b' ? rect.h - total - bottom : top);
	const clip = `${ctx.prefix}c${ctx.counter.n++}`;
	ctx.defs.push(
		`<clipPath id="${clip}"><rect x="${round(rect.x)}" y="${round(rect.y)}" width="${round(rect.w)}" height="${round(rect.h)}"/></clipPath>`,
	);
	const spans = lines.map((line, index) => {
		const x =
			line.align === 'ctr'
				? rect.x + left + innerW / 2
				: line.align === 'r'
					? rect.x + left + innerW
					: rect.x + left;
		const anchorAttr = line.align === 'ctr' ? 'middle' : line.align === 'r' ? 'end' : 'start';
		const baseline = y + (heights[index] ?? 0) * 0.8;
		y += heights[index] ?? 0;
		return `<text x="${round(x)}" y="${round(baseline)}" font-family="Calibri,Arial,sans-serif" font-size="${round(line.fontPx)}" fill="${escapeXml(line.color)}" text-anchor="${anchorAttr}"${line.bold ? ' font-weight="bold"' : ''}${line.italic ? ' font-style="italic"' : ''}>${escapeXml(line.text)}</text>`;
	});
	return `<g clip-path="url(#${escapeXml(clip)})">${spans.join('')}</g>`;
}

function markerDef(
	ctx: Context,
	end: { type?: string } | undefined,
	color: string,
): string | undefined {
	if (!end?.type || end.type === 'none') return undefined;
	const id = `${ctx.prefix}m${ctx.counter.n++}`;
	const shape =
		end.type === 'oval'
			? '<circle cx="5" cy="5" r="4"/>'
			: end.type === 'diamond'
				? '<path d="M1 5 L5 1 L9 5 L5 9 Z"/>'
				: end.type === 'arrow'
					? '<path d="M1 1 L9 5 L1 9" fill="none" stroke-width="1.5"/>'
					: '<path d="M1 1 L9 5 L1 9 Z"/>';
	ctx.defs.push(
		`<marker id="${escapeXml(id)}" viewBox="0 0 10 10" refX="5" refY="5" markerWidth="4" markerHeight="4" orient="auto-start-reverse" fill="${escapeXml(color)}" stroke="${escapeXml(color)}">${shape}</marker>`,
	);
	return id;
}

function transformOf(node: ShapeNode, box: Box): string {
	return [
		node.rotation
			? `rotate(${node.rotation} ${round(box.x + box.w / 2)} ${round(box.y + box.h / 2)})`
			: '',
		node.flipH || node.flipV
			? `translate(${node.flipH ? round(2 * box.x + box.w) : 0} ${node.flipV ? round(2 * box.y + box.h) : 0}) scale(${node.flipH ? -1 : 1} ${node.flipV ? -1 : 1})`
			: '',
	]
		.filter(Boolean)
		.join(' ');
}

function drawNode(node: ShapeNode, box: Box, ctx: Context): string {
	const transform = transformOf(node, box);
	const open = transform ? `<g transform="${escapeXml(transform)}">` : '<g>';
	if (node.kind === 'group') {
		const child = node.childFrame;
		const sx = child && child.width > 0 ? box.w / (child.width / 1) : 0;
		const sy = child && child.height > 0 ? box.h / child.height : 0;
		const inner = (node.children ?? [])
			.map((c) => {
				if (!child) return '';
				const cb: Box = {
					x: box.x + (c.frame.x - child.x) * sx,
					y: box.y + (c.frame.y - child.y) * sy,
					w: c.frame.width * sx,
					h: c.frame.height * sy,
				};
				return drawNode(c, cb, ctx);
			})
			.join('');
		return `${open}${inner}</g>`;
	}
	const id = `${ctx.prefix}g${ctx.counter.n++}`;
	const paint = shapeFill(node, id, ctx.theme);
	if (paint.defs) ctx.defs.push(paint.defs);
	const stroke = shapeStroke(node, ctx.scale, ctx.theme);
	let paths: { d: string; fill?: string | undefined; stroke?: boolean | undefined }[] = [];
	let fillNone = false;
	let textRect: { l: number; t: number; r: number; b: number } | undefined;
	if (node.paths) {
		paths = node.paths.map((path) => ({
			d: drawingPathData(path, box.w, box.h),
			fill: path.fill,
			stroke: path.stroke,
		}));
	} else if (lookupPresetShape(node.geometry)) {
		const evaluated = evaluatePresetShape(node.geometry, box.w, box.h, node.adjustments);
		if (evaluated) {
			paths = evaluated.paths;
			fillNone = evaluated.fillNone;
			textRect = evaluated.textRect;
		}
	}
	if (!paths.length) paths = [{ d: `M0 0 H${round(box.w)} V${round(box.h)} H0 Z` }];
	const head =
		node.kind === 'connector' ? markerDef(ctx, node.line?.headEnd, stroke.stroke) : undefined;
	const tail =
		node.kind === 'connector' ? markerDef(ctx, node.line?.tailEnd, stroke.stroke) : undefined;
	const body = paths
		.map((path) => {
			const fill =
				fillNone || path.fill === 'none' || node.kind === 'connector' ? 'none' : paint.fill;
			const line = path.stroke === false ? 'none' : stroke.stroke;
			return `<path d="${escapeXml(path.d)}" transform="translate(${round(box.x)} ${round(box.y)})" fill="${escapeXml(fill)}" stroke="${escapeXml(line)}" stroke-width="${round(stroke.width)}" stroke-linejoin="miter"${stroke.dash ? ` stroke-dasharray="${escapeXml(stroke.dash)}"` : ''}${head ? ` marker-start="url(#${escapeXml(head)})"` : ''}${tail ? ` marker-end="url(#${escapeXml(tail)})"` : ''}/>`;
		})
		.join('');
	const rect: Box = textRect
		? {
				x: box.x + textRect.l,
				y: box.y + textRect.t,
				w: textRect.r - textRect.l,
				h: textRect.b - textRect.t,
			}
		: box;
	return `${open}${body}${textSvg(node, rect, ctx)}</g>`;
}

/**
 * A standalone `<svg>` for a top-level shape drawn into a `widthPx` by `heightPx` box (the
 * anchor's pixel box). `scale` is the sheet zoom / 100; `idPrefix` keeps the definitions of
 * several shapes on one page apart.
 */
export function shapeSvg(
	node: ShapeNode,
	widthPx: number,
	heightPx: number,
	options: { theme?: ThemePalette | undefined; scale?: number; idPrefix?: string } = {},
): string {
	const w = Math.max(1, widthPx);
	const h = Math.max(1, heightPx);
	const ctx: Context = {
		theme: options.theme,
		scale: options.scale ?? 1,
		prefix: options.idPrefix ?? 'xs',
		counter: { n: 0 },
		defs: [],
	};
	const body = drawNode({ ...node, rotation: node.rotation ?? 0 }, { x: 0, y: 0, w, h }, ctx);
	const defs = ctx.defs.length ? `<defs>${ctx.defs.join('')}</defs>` : '';
	return `<svg xmlns="http://www.w3.org/2000/svg" width="${round(w)}" height="${round(h)}" viewBox="0 0 ${round(w)} ${round(h)}" style="overflow:visible" role="img" aria-label="${escapeXml(node.name ?? 'Shape')}">${defs}${body}</svg>`;
}
