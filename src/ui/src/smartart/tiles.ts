/**
 * SmartArt gallery tiles every product draws: a small diagram per layout family (list, process,
 * cycle, hierarchy, ...) in Accent 1, and a three-box process filled in a colour scheme's accents
 * or in Accent 1 with a style's effect, both on a white card. SVG markup only, from trusted
 * geometry and sanitised colours. Moved from the PowerPoint UI
 * (`pptx/render/ribbon-galleries/smartart-layout-tiles.ts` and `smartart-gallery-tiles.ts`).
 */
import type { DiagramStyleIntensity } from 'ooxml-core/diagram';
import { safeColor, svgTile } from '../ribbon/svg-preview';

export interface SmartArtTileSize {
	width: number;
	height: number;
}

const round = (value: number): string => String(Math.round(value * 100) / 100);
type Draw = (fill: string, w: number, h: number) => string;

const rect = (x: number, y: number, w: number, h: number, fill: string, rx = 1.5): string =>
	`<rect x="${round(x)}" y="${round(y)}" width="${round(w)}" height="${round(h)}" rx="${rx}" fill="${fill}"/>`;
const circle = (x: number, y: number, r: number, fill: string, opacity = 1): string =>
	`<circle cx="${round(x)}" cy="${round(y)}" r="${round(r)}" fill="${fill}" fill-opacity="${opacity}"/>`;
const poly = (points: ReadonlyArray<readonly [number, number]>, fill: string): string =>
	`<polygon points="${points.map(([x, y]) => `${round(x)},${round(y)}`).join(' ')}" fill="${fill}"/>`;
const line = (x1: number, y1: number, x2: number, y2: number, stroke: string): string =>
	`<path d="M${round(x1)} ${round(y1)}L${round(x2)} ${round(y2)}" stroke="${stroke}" stroke-width="1" fill="none"/>`;

const DRAWINGS: Readonly<Record<string, Draw>> = {
	list: (f, w, h) =>
		[0, 1, 2].map((i) => rect(6, 5 + i * (h - 10) * 0.36, w - 12, (h - 10) * 0.26, f)).join(''),
	process: (f, w, h) =>
		[0, 1, 2]
			.map((i) => rect(4 + i * ((w - 8) / 3), h * 0.3, (w - 8) / 3 - 4, h * 0.4, f))
			.join(''),
	cycle: (f, w, h) =>
		(
			[
				[w / 2, 8],
				[w - 14, h / 2],
				[w / 2, h - 8],
				[14, h / 2],
			] as const
		)
			.map(([x, y]) => circle(x, y, 5, f))
			.join(''),
	hierarchy: (f, w) =>
		rect(w / 2 - 9, 4, 18, 8, f) +
		line(w / 2, 12, w / 2, 18, f) +
		line(14, 18, w - 14, 18, f) +
		rect(6, 19, 18, 8, f) +
		rect(w - 24, 19, 18, 8, f),
	relationship: (f, w, h) =>
		circle(w * 0.32, h / 2, 9, f) +
		circle(w * 0.68, h / 2, 9, f) +
		line(w * 0.32 + 9, h / 2, w * 0.68 - 9, h / 2, '#FFFFFF'),
	matrix: (f, w) =>
		[0, 1, 2, 3]
			.map((i) =>
				rect(10 + (i % 2) * ((w - 20) / 2), 5 + Math.floor(i / 2) * 14, (w - 20) / 2 - 3, 11, f),
			)
			.join(''),
	pyramid: (f, w, h) =>
		poly(
			[
				[w / 2, 4],
				[w / 2 + 6, 13],
				[w / 2 - 6, 13],
			],
			f,
		) +
		poly(
			[
				[w / 2 - 7, 15],
				[w / 2 + 7, 15],
				[w / 2 + 13, 24],
				[w / 2 - 13, 24],
			],
			f,
		) +
		poly(
			[
				[w / 2 - 14, 26],
				[w / 2 + 14, 26],
				[w / 2 + 20, h - 4],
				[w / 2 - 20, h - 4],
			],
			f,
		),
	funnel: (f, w) =>
		[0, 1, 2]
			.map((i) => {
				const half = 20 - i * 6;
				return poly(
					[
						[w / 2 - half, 5 + i * 9],
						[w / 2 + half, 5 + i * 9],
						[w / 2 + half - 4, 12 + i * 9],
						[w / 2 - half + 4, 12 + i * 9],
					],
					f,
				);
			})
			.join(''),
	target: (f, w, h) =>
		circle(w / 2, h / 2, 14, f, 0.35) +
		circle(w / 2, h / 2, 9, f, 0.6) +
		circle(w / 2, h / 2, 4, f),
	gear: (f, w, h) =>
		circle(w * 0.38, h * 0.55, 9, f) +
		circle(w * 0.38, h * 0.55, 3, '#FFFFFF') +
		circle(w * 0.66, h * 0.36, 6, f) +
		circle(w * 0.66, h * 0.36, 2, '#FFFFFF'),
	venn: (f, w, h) =>
		circle(w * 0.38, h * 0.42, 10, f, 0.55) +
		circle(w * 0.62, h * 0.42, 10, f, 0.55) +
		circle(w * 0.5, h * 0.64, 10, f, 0.55),
	timeline: (f, w, h) =>
		line(5, h / 2, w - 5, h / 2, f) +
		[0, 1, 2].map((i) => circle(12 + i * ((w - 24) / 2), h / 2, 3.5, f)).join('') +
		[0, 1, 2].map((i) => rect(7 + i * ((w - 24) / 2), i % 2 ? h / 2 + 6 : 5, 10, 7, f)).join(''),
	chevron: (f, w, h) =>
		[0, 1, 2]
			.map((i) => {
				const x = 4 + i * ((w - 8) / 3);
				const cw = (w - 8) / 3;
				return poly(
					[
						[x, h * 0.3],
						[x + cw - 5, h * 0.3],
						[x + cw, h / 2],
						[x + cw - 5, h * 0.7],
						[x, h * 0.7],
						[x + 5, h / 2],
					],
					f,
				);
			})
			.join(''),
	bending: (f, w, h) =>
		rect(5, 5, 14, 9, f) +
		rect(w / 2 - 7, 5, 14, 9, f) +
		rect(w - 19, 5, 14, 9, f) +
		rect(w - 19, h - 14, 14, 9, f) +
		rect(w / 2 - 7, h - 14, 14, 9, f),
};

/** A layout family's tile; unknown families fall back to a row of three boxes. */
export function smartArtLayoutTileSvg(
	type: string,
	fill: string | undefined,
	size: SmartArtTileSize,
): string {
	const color = safeColor(fill, '#4472C4');
	const draw = DRAWINGS[type] ?? DRAWINGS.process!;
	const bg = rect(0, 0, size.width, size.height, '#FFFFFF', 0);
	return svgTile(size.width, size.height, '', bg + draw(color, size.width, size.height));
}

/** Three boxes with arrow gaps, filled with `fills` (cycled) and an optional effect. */
export function smartArtTileSvg(
	id: string,
	fills: readonly string[],
	style: DiagramStyleIntensity,
	size: SmartArtTileSize,
): string {
	const safeId = `sat-${id}`.replace(/[^\w-]/gu, '_');
	const pad = 4;
	const gap = 4;
	const w = (size.width - pad * 2 - gap * 2) / 3;
	const h = size.height * 0.46;
	const y = (size.height - h) / 2;
	let defs = '';
	let filter = '';
	if (style !== 'flat') {
		const intense = style === 'intense';
		defs += `<filter id="${safeId}-s" x="-30%" y="-30%" width="160%" height="170%"><feDropShadow dx="0" dy="${intense ? 1.6 : 0.9}" stdDeviation="${intense ? 1.2 : 0.7}" flood-color="#000000" flood-opacity="${intense ? 0.45 : 0.3}"/></filter>`;
		filter = ` filter="url(#${safeId}-s)"`;
	}
	const boxes: string[] = [];
	for (let i = 0; i < 3; i++) {
		const color = safeColor(fills[i % fills.length], '#4472C4');
		let paint = color;
		if (style === 'intense') {
			defs += `<linearGradient id="${safeId}-g${i}" x1="0" y1="0" x2="0" y2="1"><stop offset="0" stop-color="${color}" stop-opacity="0.8"/><stop offset="1" stop-color="${color}"/></linearGradient>`;
			paint = `url(#${safeId}-g${i})`;
		}
		const x = pad + i * (w + gap);
		boxes.push(
			`<rect x="${round(x)}" y="${round(y)}" width="${round(w)}" height="${round(h)}" rx="${round(Math.min(w, h) * 0.12)}" fill="${paint}" stroke="#FFFFFF" stroke-width="0.75"${filter}/>`,
		);
	}
	const bg = `<rect x="0" y="0" width="${size.width}" height="${size.height}" fill="#FFFFFF"/>`;
	return svgTile(size.width, size.height, defs, bg + boxes.join(''));
}
