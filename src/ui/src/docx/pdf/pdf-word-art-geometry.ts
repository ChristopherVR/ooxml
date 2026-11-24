/**
 * Where WordArt's letters go in the Word PDF export: each line of `wordArtLayout` centred on its
 * warp path (or straight baseline), the point and direction under each letter, the box its
 * gradient spans, and the gradient itself. Shared by the outline and standard-font painters of
 * `pdf-word-art.ts`. DOM-free.
 */
import type { TextBoxContent } from 'ooxml-core/docx';
import type { WordArtLayout, WordArtLayoutLine } from 'ooxml-core/drawingml';
import type { GlyphPointMap } from '../../export/pdf-glyph-outlines';
import { parsePath, type PathSegment } from '../../export/pdf-path';
import { pathSampler, type PathPoint } from '../../export/pdf-path-sampler';
import { pathBounds } from '../../export/pdf-svg-gradient';
import type { PdfVectorPage } from '../../export/pdf-vector';
import { cssRgba } from './pdf-text';

/** Arial's line box, for the gradient's bounding box when no font file gives it. */
export const ASCENT = 0.905;
export const DESCENT = 0.212;

export interface Slot {
	char: string;
	/** What the standard font writes for it. */
	code: string;
	glyph: number;
	x: number;
	advance: number;
}

export interface LinePlacement {
	map: GlyphPointMap;
	/** The point (and direction) at `u` along the line, or `undefined` past a path's ends. */
	frame(u: number): PathPoint | undefined;
	bend: boolean;
}

/** The lines Print Layout gives WordArt: each paragraph's runs joined. */
export function wordArtLines(content: TextBoxContent): string[] {
	return content.blocks.flatMap((block) =>
		block.type === 'paragraph' ? [block.runs.map((run) => run.text).join('')] : [],
	);
}

export function placement(line: WordArtLayoutLine, total: number): LinePlacement | undefined {
	if (line.kind === 'straight') {
		const left = line.x - total / 2;
		return {
			map: (u, v) => [left + u, line.y + v],
			frame: (u) => ({ x: left + u, y: line.y, angle: 0 }),
			bend: false,
		};
	}
	const segments = parsePath(line.path);
	const sampler = segments && pathSampler(segments);
	if (!sampler) return undefined;
	// `startOffset="50%"` with `text-anchor="middle"`: the line is centred on the path.
	const start = sampler.length / 2 - total / 2;
	const at = (u: number) => sampler.at(start + u);
	return {
		map: (u, v) => {
			const p = at(u);
			return [
				line.offsetX + p.x - Math.sin(p.angle) * v,
				line.offsetY + p.y + Math.cos(p.angle) * v,
			];
		},
		frame: (u) => {
			const distance = start + u;
			if (distance < 0 || distance > sampler.length) return undefined;
			const p = at(u);
			return { x: line.offsetX + p.x, y: line.offsetY + p.y, angle: p.angle };
		},
		bend: true,
	};
}

/** The box `objectBoundingBox` gradients map to: every visible glyph cell, as placed. */
export function cellBounds(
	slots: readonly Slot[],
	place: LinePlacement,
	ascent: number,
	descent: number,
): [number, number, number, number] | undefined {
	const corners: PathSegment[] = [];
	for (const slot of slots) {
		if (!place.frame(slot.x + slot.advance / 2)) continue;
		for (const u of [slot.x, slot.x + slot.advance])
			for (const v of [-ascent, descent]) {
				const [x, y] = place.map(u, v);
				corners.push({ t: 'L', x, y });
			}
	}
	return pathBounds(corners);
}

export const rgb = (hex: string) => cssRgba(hex) ?? { r: 0, g: 0, b: 0, a: 1 };

/** Paints the fill (flat or the gradient over `bounds`) inside the current clip. */
export function paintGradient(
	page: PdfVectorPage,
	layout: WordArtLayout,
	bounds: [number, number, number, number],
): void {
	const gradient = layout.gradient!;
	const angle = (gradient.angle / 10800000) * Math.PI;
	const dx = Math.cos(angle) / 2;
	const dy = Math.sin(angle) / 2;
	const from = rgb(layout.fill);
	const to = rgb(gradient.end);
	page.concat([bounds[2] || 1, 0, 0, bounds[3] || 1, bounds[0], bounds[1]]);
	page.paintShading({
		kind: 'axial',
		coords: [0.5 - dx, 0.5 - dy, 0.5 + dx, 0.5 + dy],
		stops: [
			{ offset: 0, r: from.r, g: from.g, b: from.b },
			{ offset: 1, r: to.r, g: to.g, b: to.b },
		],
	});
}
