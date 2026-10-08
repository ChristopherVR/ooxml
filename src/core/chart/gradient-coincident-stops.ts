/**
 * Native Excel paints two stops saved at the same position as a one-step ramp
 * that starts at the shared position, not as a hard edge. Measured on the
 * `coincident` profile of `__fixtures__/native-gradient-linear-profiles.json`
 * (Excel 16.0 build 20430): pixels whose centre projects to exactly the shared
 * position keep the earlier stop, and pixels 0.000469 past it already carry
 * 30/255 of the later stop, which bounds the ramp width to 1/261..1/251 of the
 * gradient vector. Only linear fills were measured.
 */
export const NATIVE_COINCIDENT_STOP_RAMP = 1 / 256;

/** Separate coincident sorted stops by the native ramp, without crossing the next stop. */
export function spreadCoincidentStops<
	T extends { offset: number; color: string; opacity?: number },
>(stops: readonly T[]): T[] {
	const out: T[] = [];
	for (const [index, stop] of stops.entries()) {
		const source = stops[index - 1];
		const previous = out.at(-1);
		const next = stops[index + 1];
		if (
			!source ||
			!previous ||
			stop.offset !== source.offset ||
			(stop.color === source.color && (stop.opacity ?? 1) === (source.opacity ?? 1))
		) {
			out.push(stop);
			continue;
		}
		const limit = Math.min(1, next ? next.offset : 1);
		out.push({
			...stop,
			offset: Math.max(stop.offset, Math.min(previous.offset + NATIVE_COINCIDENT_STOP_RAMP, limit)),
		});
	}
	return out;
}
