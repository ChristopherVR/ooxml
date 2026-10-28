/** Pixel evidence over a white background, without registration or content cropping. */
export function comparePixels(native: Uint8ClampedArray, shared: Uint8ClampedArray) {
	if (native.length !== shared.length || native.length === 0 || native.length % 4)
		throw new Error('Comparable non-empty RGBA buffers are required');
	let absolute = 0;
	let changed = 0;
	let nativeInk = 0;
	let sharedInk = 0;
	let intersection = 0;
	let union = 0;
	const difference = new Uint8ClampedArray(native.length);
	for (let offset = 0; offset < native.length; offset += 4) {
		let maximum = 0;
		let nativeDark = false;
		let sharedDark = false;
		for (let channel = 0; channel < 3; channel++) {
			const a = 255 + ((native[offset + channel]! - 255) * native[offset + 3]!) / 255;
			const b = 255 + ((shared[offset + channel]! - 255) * shared[offset + 3]!) / 255;
			const delta = Math.abs(a - b);
			absolute += delta;
			maximum = Math.max(maximum, delta);
			nativeDark ||= a < 245;
			sharedDark ||= b < 245;
			difference[offset + channel] = Math.min(255, delta * 4);
		}
		difference[offset + 3] = 255;
		if (maximum > 16) changed++;
		if (nativeDark) nativeInk++;
		if (sharedDark) sharedInk++;
		if (nativeDark && sharedDark) intersection++;
		if (nativeDark || sharedDark) union++;
	}
	const pixels = native.length / 4;
	return {
		normalizedMeanAbsoluteError: absolute / (pixels * 3 * 255),
		changedPixelRatio: changed / pixels,
		foregroundIntersectionOverUnion: union ? intersection / union : 0,
		nativeForegroundRatio: nativeInk / pixels,
		sharedForegroundRatio: sharedInk / pixels,
		difference,
	};
}
