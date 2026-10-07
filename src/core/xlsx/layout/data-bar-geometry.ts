interface BarGeometryOptions {
	axis: 'automatic' | 'middle' | 'none';
	autoMin: boolean;
	autoMax: boolean;
	minLength: number;
	maxLength: number;
}

/** Normalized geometry. Printing insets and pixel snapping belong to the renderer. */
export function dataBarGeometry(
	value: number,
	low: number,
	high: number,
	options: BarGeometryOptions,
) {
	const lo = options.autoMin ? Math.min(0, low) : low;
	const hi = options.autoMax ? Math.max(0, high) : high;
	const clamp = (v: number) => Math.max(0, Math.min(1, v));
	if (options.axis === 'none') {
		const ratio = hi > lo ? clamp((value - lo) / (hi - lo)) : 0.5;
		return {
			start: 0,
			fraction: clamp(options.minLength + ratio * (options.maxLength - options.minLength)),
			reverse: false,
		};
	}
	if (options.axis === 'middle') {
		const span = 2 * Math.max(Math.abs(lo), Math.abs(hi));
		const fraction = span > 0 ? Math.min(0.5, Math.abs(value) / span) : 0;
		return { start: value < 0 ? 0.5 - fraction : 0.5, fraction, axis: 0.5, reverse: value < 0 };
	}
	if (lo === 0 && hi === 0) return { start: 0.5, fraction: 0, axis: 0.5, reverse: false };
	if (lo >= 0) {
		const fraction = hi > lo ? clamp((value - lo) / (hi - lo)) : 0.5;
		return { start: 0, fraction, reverse: false };
	}
	if (hi <= 0) {
		const fraction = hi > lo ? clamp((hi - value) / (hi - lo)) : 0.5;
		return { start: 1 - fraction, fraction, reverse: true };
	}
	const axis = -lo / (hi - lo);
	const bounded = Math.max(lo, Math.min(hi, value));
	const fraction = Math.abs(bounded) / (hi - lo);
	return { start: bounded < 0 ? axis - fraction : axis, fraction, axis, reverse: bounded < 0 };
}
