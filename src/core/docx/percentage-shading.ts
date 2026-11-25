import { isStShd, type StShd } from './generated/wml-simple-types';
import { parseRgbColor } from './simple-types';
import { getW, type XmlElement } from './xml';

export interface PercentageShading {
	pattern: Extract<StShd, `pct${number}`>;
	/** Foreground colour. Word uses black for auto. */
	color?: string;
}

function isPercentage(value: string): value is PercentageShading['pattern'] {
	return value.startsWith('pct') && isStShd(value);
}

export function isPercentageShading(value: unknown): value is PercentageShading {
	return (
		typeof value === 'object' &&
		value !== null &&
		'pattern' in value &&
		typeof value.pattern === 'string' &&
		isPercentage(value.pattern) &&
		(!('color' in value) || typeof value.color === 'string')
	);
}

export function parsePercentageShading(
	element: XmlElement | undefined,
): PercentageShading | undefined {
	const pattern = getW(element, 'val');
	if (!pattern || !isPercentage(pattern)) return undefined;
	const color = parseRgbColor(getW(element, 'color'));
	return { pattern, ...(color ? { color } : {}) };
}

/** Word paints percentage shades as a solid blend of the foreground and background. */
export function percentageShadingColor(
	fill: string | undefined,
	shading: PercentageShading | undefined,
): string | undefined {
	if (!shading) return fill;
	const rgb = (color: string | undefined, fallback: number) =>
		color && /^#?[0-9a-f]{6}$/i.test(color)
			? Number.parseInt(color.replace(/^#/, ''), 16)
			: fallback;
	const foreground = rgb(shading.color, 0);
	const background = rgb(fill, 0xffffff);
	const amount = Number(shading.pattern.slice(3)) / 100;
	const hex = [16, 8, 0].map((shift) => {
		const from = (background >> shift) & 255;
		const to = (foreground >> shift) & 255;
		return Math.floor(from + (to - from) * amount)
			.toString(16)
			.padStart(2, '0');
	});
	return `#${hex.join('')}`;
}
