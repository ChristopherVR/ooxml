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
	const minimum = options.minLength;
	const delta = options.maxLength - minimum;
	const interpolate = (ratio: number, offset = minimum) => clamp(offset + clamp(ratio) * delta);
	if (options.axis === 'none') {
		const ratio = hi > lo ? clamp((value - lo) / (hi - lo)) : 0.5;
		return {
			start: 0,
			fraction: interpolate(ratio),
			reverse: false,
		};
	}
	if (options.axis === 'middle') {
		const span = Math.max(Math.abs(lo), Math.abs(hi));
		const fraction = interpolate(span > 0 ? Math.abs(value) / span : 0) / 2;
		return { start: value < 0 ? 0.5 - fraction : 0.5, fraction, axis: 0.5, reverse: value < 0 };
	}
	if (lo === 0 && hi === 0) return { start: 0.5, fraction: minimum / 2, axis: 0.5, reverse: false };
	if (lo >= 0) {
		const fraction =
			hi > lo
				? interpolate((value - lo) / (hi - lo), lo === 0 ? minimum / 2 : minimum)
				: interpolate(0.5);
		return { start: 0, fraction, reverse: false };
	}
	if (hi <= 0) {
		const fraction =
			hi > lo
				? interpolate((hi - value) / (hi - lo), hi === 0 ? minimum / 2 : minimum)
				: interpolate(0.5);
		const axis = hi === 0 ? (1 + options.maxLength - minimum) / 2 : options.maxLength;
		return {
			start: axis - fraction,
			fraction,
			...(axis > 0 && axis < 1 ? { axis } : {}),
			reverse: true,
		};
	}
	const axis = (1 - options.maxLength + minimum) / 2 + (-lo / (hi - lo)) * delta;
	const bounded = Math.max(lo, Math.min(hi, value));
	const fraction = interpolate(Math.abs(bounded) / (hi - lo), minimum / 2);
	return { start: bounded < 0 ? axis - fraction : axis, fraction, axis, reverse: bounded < 0 };
}
