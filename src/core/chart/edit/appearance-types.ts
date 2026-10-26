// Resolved pane/host appearance values. Units are points/degrees and colours are hex; these
// contracts do not import UI or product types, and are shared by the Office adapters.
import type { ChartGradientFill } from '../gradient-definition';

export interface ChartShadowAppearance {
	color: string;
	transparency?: number;
	blur: number;
	distance: number;
	angle: number;
}
export interface ChartGlowAppearance {
	color: string;
	transparency?: number;
	radius: number;
}
export interface ChartEffectsAppearance {
	shadow?: ChartShadowAppearance | null;
	glow?: ChartGlowAppearance | null;
	softEdge?: number | null;
}
export interface ChartThreeDAppearance {
	bevelTop?: { type: string; width: number; height: number } | null;
	bevelBottom?: { type: string; width: number; height: number } | null;
	depth?: { size: number; color?: string } | null;
	contour?: { size: number; color?: string } | null;
	material?: string | null;
	lighting?: string | null;
}
export interface ChartFillAppearance {
	kind: 'none' | 'solid' | 'gradient' | 'pattern' | 'picture' | 'automatic';
	color?: string;
	transparency?: number;
	gradient?: ChartGradientFill;
	pattern?: { preset: string; foreground: string; background: string };
	/** The host allocates relationships before conversion. Raw image data is refused. */
	picture?: {
		relId?: string;
		linkRelId?: string;
		src?: string;
		tile?: boolean;
		transparency?: number;
	};
}
export interface ChartLineAppearance {
	kind: 'none' | 'solid' | 'gradient' | 'automatic';
	color?: string;
	transparency?: number;
	gradient?: ChartGradientFill;
	width?: number;
	dash?: string;
	compound?: string;
	cap?: string;
	join?: string;
	beginArrow?: string;
	endArrow?: string;
}
export interface ChartTextAppearance {
	typeface?: string | null;
	eastAsiaTypeface?: string | null;
	sizePt?: number | null;
	bold?: boolean;
	italic?: boolean;
	fill?: Pick<ChartFillAppearance, 'kind' | 'color' | 'transparency' | 'gradient'>;
	outline?: Pick<ChartLineAppearance, 'kind' | 'color' | 'transparency' | 'width' | 'dash'>;
	effects?: Pick<ChartEffectsAppearance, 'shadow' | 'glow'>;
	rotation?: number | null;
	direction?: string | null;
	anchor?: string | null;
}
