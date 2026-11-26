import { NS, first, type XmlElement } from '../xml/index';
import { parseDrawingColorIn } from './drawing-color';
import type { DrawingColor } from './types';
import { drawingColorXml } from './write-color';

export interface DrawingStyleRef {
	idx: number;
	color?: DrawingColor;
}

/** The four references shared by DrawingML shapes and pictures. */
export interface DrawingShapeStyleRefs {
	line?: DrawingStyleRef;
	fill?: DrawingStyleRef;
	effect?: DrawingStyleRef;
	font?: { idx?: string; color?: DrawingColor };
}

export function parseDrawingShapeStyle(
	node: XmlElement | undefined,
): DrawingShapeStyleRefs | undefined {
	if (!node) return undefined;
	const ref = (tag: string): DrawingStyleRef | undefined => {
		const child = first(node, tag, NS.a);
		if (!child) return undefined;
		const idx = Number(child.getAttribute('idx') ?? 0);
		const color = parseDrawingColorIn(child);
		return { idx: Number.isFinite(idx) ? idx : 0, ...(color ? { color } : {}) };
	};
	const line = ref('lnRef');
	const fill = ref('fillRef');
	const effect = ref('effectRef');
	const font = first(node, 'fontRef', NS.a);
	const fontIndex = font?.getAttribute('idx');
	const fontColor = parseDrawingColorIn(font);
	return {
		...(line ? { line } : {}),
		...(fill ? { fill } : {}),
		...(effect ? { effect } : {}),
		...(font
			? {
					font: {
						...(fontIndex ? { idx: fontIndex } : {}),
						...(fontColor ? { color: fontColor } : {}),
					},
				}
			: {}),
	};
}

export type CompleteStyleRefs = Required<
	Pick<DrawingShapeStyleRefs, 'line' | 'fill' | 'effect'>
> & {
	font: { idx: 'major' | 'minor' | 'none'; color: DrawingColor };
};

export function drawingShapeStyleRefsXml(style: CompleteStyleRefs): string {
	const ref = (tag: string, idx: number | string, color: DrawingColor | undefined) =>
		`<a:${tag} idx="${idx}">${color ? drawingColorXml(color) : ''}</a:${tag}>`;
	return (
		ref('lnRef', style.line.idx, style.line.color) +
		ref('fillRef', style.fill.idx, style.fill.color) +
		ref('effectRef', style.effect.idx, style.effect.color) +
		ref('fontRef', style.font.idx, style.font.color)
	);
}
