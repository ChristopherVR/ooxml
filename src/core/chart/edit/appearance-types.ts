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
	compound?: string | undefined;
	cap?: string | undefined;
	join?: string;
	beginArrow?: string;
	endArrow?: string;
}
export interface ChartTextAppearance {
	typeface?: string | null | undefined;
	eastAsiaTypeface?: string | null | undefined;
	sizePt?: number | null | undefined;
	bold?: boolean | undefined;
	italic?: boolean | undefined;
	fill?: Pick<ChartFillAppearance, 'kind' | 'color' | 'transparency' | 'gradient'> | undefined;
	outline?:
		| Pick<ChartLineAppearance, 'kind' | 'color' | 'transparency' | 'width' | 'dash'>
		| undefined;
	effects?: Pick<ChartEffectsAppearance, 'shadow' | 'glow'> | undefined;
	rotation?: number | null | undefined;
	direction?: string | null | undefined;
	anchor?: string | null | undefined;
	/** True wraps text in its box, false never wraps; null restores the default. */
	wrap?: boolean | null | undefined;
}
