import type { DiagramColor } from './types';

/** Native Office brightness, in percent. Noncanonical luminance expressions remain unknown. */
export function drawingColorBrightness(color: DiagramColor): number | undefined {
	const transforms = color.transforms.filter((item) =>
		['lum', 'lumMod', 'lumOff'].includes(item.name),
	);
	if (!transforms.length) return 0;
	const mod = transforms.find((item) => item.name === 'lumMod');
	const off = transforms.find((item) => item.name === 'lumOff');
	if (!mod || transforms[0] !== mod || transforms.length !== (off ? 2 : 1)) return undefined;
	const multiplier = Number(mod.value);
	const offset = off ? Number(off.value) : 0;
	if (
		!mod.value ||
		(off && !off.value) ||
		!Number.isFinite(multiplier) ||
		!Number.isFinite(offset) ||
		multiplier < 0 ||
		multiplier > 100000 ||
		offset < 0 ||
		offset > 100000
	)
		return undefined;
	if (off && multiplier + offset !== 100000) return undefined;
	return Math.round(off ? offset : multiplier - 100000) / 1000;
}

/** Replace native brightness while preserving the color choice and unrelated transforms. */
export function withDrawingColorBrightness(color: DiagramColor, percent: number): DiagramColor {
	if (!Number.isFinite(percent) || percent < -100 || percent > 100)
		throw new RangeError('Brightness must be from -100 to 100');
	const result = structuredClone(color);
	result.transforms = result.transforms.filter((item) => !['lumMod', 'lumOff'].includes(item.name));
	result.transforms.push({
		name: 'lumMod',
		value: String(Math.round((100 - Math.abs(percent)) * 1000)),
	});
	if (percent > 0)
		result.transforms.push({ name: 'lumOff', value: String(Math.round(percent * 1000)) });
	return result;
}
