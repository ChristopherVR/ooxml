import type { DrawingEffect, DrawingOuterShadow } from './effect-model';
import type { ThemeEffectStyle } from './format-scheme';
import type { DrawingColor } from './types';

const px = (value: number) => Math.round(value * 9525);
const black = (opacity: number): DrawingColor => ({
	kind: 'srgb',
	value: '000000',
	transforms: [{ name: 'alpha', value: String(opacity * 1000) }],
});
const accent: DrawingColor = { kind: 'scheme', value: 'phClr', transforms: [] };
const outer = (
	blur: number,
	distance: number,
	opacity: number,
	angle = 90,
): DrawingOuterShadow => ({
	kind: 'outerShadow',
	blurRadiusEmu: px(blur),
	distanceEmu: px(distance),
	directionAngle: angle * 60000,
	color: black(opacity),
});
const inner = (blur: number, distance: number, opacity: number): DrawingEffect => ({
	...outer(blur, distance, opacity),
	kind: 'innerShadow',
});
const glow = (radius: number): DrawingEffect => ({
	kind: 'glow',
	radiusEmu: px(radius),
	color: accent,
});
const edge = (radius: number): DrawingEffect => ({ kind: 'softEdge', radiusEmu: px(radius) });
const reflection: DrawingEffect = {
	kind: 'reflection',
	blurRadiusEmu: px(1),
	distanceEmu: px(1),
	directionAngle: 5400000,
	fadeDirectionAngle: 5400000,
	startAlpha: 45000,
	startPosition: 0,
	endAlpha: 0,
	endPosition: 55000,
	scaleY: -100000,
	alignment: 'bl',
};

function preset<Name extends string>(
	name: Name,
	subtle: DrawingEffect[],
	moderate: DrawingEffect[],
	intense: DrawingEffect[],
) {
	const styles: ThemeEffectStyle[] = [subtle, moderate, intense].map((effects) => ({
		effects: { effects },
	}));
	return { name, styles };
}

// Authored DrawingML matrices for familiar gallery categories. The tiers are ordered by style index.
export const THEME_EFFECT_PRESETS = [
	preset('Office', [], [outer(3, 1.5, 24)], [outer(4, 2.5, 30)]),
	preset('Office 2007-2010', [], [outer(4, 2, 32, 45)], [outer(5, 4, 42, 45)]),
	preset('Subtle Solid', [], [], [outer(2, 1, 18)]),
	preset('Banded Edge', [], [inner(1, 1, 30)], [inner(2, 2, 45)]),
	preset('Smokey Glass', [], [edge(1), outer(3, 2, 32)], [edge(3), outer(5, 3, 48)]),
	preset('Glow Edge', [], [glow(3)], [glow(5), outer(3, 2, 24)]),
	preset('Grunge Texture', [], [edge(2)], [edge(3), inner(2, 2, 40)]),
	preset('Frosted Glass', [], [edge(1), glow(2)], [edge(2), glow(4)]),
	preset('Top Shadow', [], [outer(3, 3, 40, 270)], [outer(6, 5, 48, 270)]),
	preset('Inset', [], [inner(3, 2, 40)], [inner(5, 3, 60)]),
	preset('Milk Glass', [], [inner(2, 1, 18)], [reflection]),
	preset('Riblet', [], [inner(2, 2, 38)], [inner(3, 3, 50), outer(3, 2, 24)]),
	preset('Ripple', [], [reflection], [reflection, glow(3)]),
	preset('Slanted', [], [outer(3, 4, 38, 35)], [outer(5, 6, 40, 35)]),
	preset('Glossy', [], [inner(2, 1, 22)], [inner(3, 2, 32), reflection]),
];

export type ThemeEffectPresetName = (typeof THEME_EFFECT_PRESETS)[number]['name'];

export function sameThemeEffectStyles(a: ThemeEffectStyle[], b: ThemeEffectStyle[]): boolean {
	const signature = (styles: ThemeEffectStyle[]) =>
		JSON.stringify(styles, (key, value: unknown) => {
			if (key === 'sourceXml') return undefined;
			if (typeof value !== 'object' || value === null || Array.isArray(value)) return value;
			return Object.fromEntries(
				Object.entries(value).sort(([left], [right]) => (left < right ? -1 : left > right ? 1 : 0)),
			);
		});
	return signature(a) === signature(b);
}
