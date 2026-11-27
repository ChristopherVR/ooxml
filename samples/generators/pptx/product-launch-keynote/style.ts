/**
 * The design system of the Aerie keynote: one palette, one type scale, one grid. Every slide
 * module reads its colours, sizes and positions from here.
 */
import type {
	FillInput,
	ShadowInput,
	TextStyleInput,
} from '../../../../src/core/pptx/core/builders/sdk/index';
import { FONTS } from '../../lib/fiction';
import { pt } from '../kit/geometry';

/** Night sky, two accents (dawn amber and sky cyan); the violet only appears inside gradients. */
export const C = {
	night: '#070B14',
	deep: '#0E1626',
	panel: '#141C2E',
	panelHi: '#1B2640',
	line: '#2A3654',
	text: '#F2F5FA',
	soft: '#C9D3E6',
	muted: '#8D9AB3',
	amber: '#FFB347',
	sky: '#4CC9F0',
	violet: '#7B61FF',
	coral: '#FF7A59',
} as const;

/** Type scale in points. */
export const T = {
	hero: 72,
	display: 54,
	title: 34,
	stat: 48,
	lead: 20,
	body: 15,
	small: 12,
	caption: 10.5,
} as const;

/** Grid: 80 px side margins, a 1120 px measure, twelve columns with a 24 px gutter. */
export const G = {
	left: 80,
	right: 1200,
	width: 1120,
	top: 64,
	/** Where content starts below a kicker and a one-line title. */
	content: 200,
	bottom: 640,
	gutter: 24,
} as const;

/** The x of column `index` (0 to 11) and the width of `span` columns. */
export function col(index: number, span: number): { x: number; width: number } {
	const unit = (G.width - G.gutter * 11) / 12;
	return { x: G.left + index * (unit + G.gutter), width: span * unit + (span - 1) * G.gutter };
}

export const display = (size: number, color: string = C.text): Partial<TextStyleInput> => ({
	fontFamily: FONTS.display,
	fontSize: pt(size),
	bold: true,
	color,
});

export const body = (size: number = T.body, color: string = C.soft): Partial<TextStyleInput> => ({
	fontFamily: FONTS.body,
	fontSize: pt(size),
	color,
});

/** Small tracked capitals, used for labels above numbers and cards. */
export const label = (color: string = C.muted): Partial<TextStyleInput> => ({
	fontFamily: FONTS.display,
	fontSize: pt(T.caption),
	bold: true,
	color,
	characterSpacing: 2,
});

/** A card surface: a faint top-to-bottom gradient on the panel colour. */
export const cardFill: FillInput = {
	type: 'gradient',
	angle: 90,
	stops: [
		{ color: C.panelHi, position: 0 },
		{ color: C.panel, position: 100 },
	],
};

/** The one soft shadow cards and pictures share. */
export const softShadow: ShadowInput = {
	color: '#000000',
	blur: 36,
	offsetY: 14,
	opacity: 0.45,
};

/** A round glow: the colour fading to nothing at the rim. */
export function glowFill(color: string, strength = 0.55): FillInput {
	return {
		type: 'gradient',
		gradientType: 'radial',
		path: 'circle',
		focus: { l: 0.5, t: 0.5, r: 0.5, b: 0.5 },
		stops: [
			// A circle path gradient ends at the corners of the box, so the rim of the ellipse sits
			// at about 70 %: the colour has to be gone by then or the edge shows.
			{ color, position: 0, opacity: strength },
			{ color, position: 30, opacity: strength * 0.4 },
			{ color, position: 68, opacity: 0 },
			{ color, position: 100, opacity: 0 },
		],
	};
}

export const noLine = { color: 'transparent' } as const;
