import { lerpColor } from './data';
export function sequentialColorScale(t: number, accent?: string): string {
	const clamped = Math.max(0, Math.min(1, t));
	if (accent && /^#[0-9a-f]{6}$/iu.test(accent))
		return lerpColor(lerpColor('#ffffff', accent, 0.2), accent, clamped);
	if (clamped <= 0.5) {
		return lerpColor('#dbeafe', '#3b82f6', clamped * 2);
	}
	return lerpColor('#3b82f6', '#1e3a5f', (clamped - 0.5) * 2);
}

/** Normalise a value to [0..1] within a min/max range. */
export function normalizeValue(value: number, min: number, max: number): number {
	if (max === min) {
		return 0.5;
	}
	return (value - min) / (max - min);
}
