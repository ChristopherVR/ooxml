import type { InkXmlObject } from './types';
import { isFiniteNumber } from '../../units/finite';
export { decodeTracePoints } from './inkml-trace-points';

/** Channel names in trace-column order (upper-cased), e.g. ["X", "Y", "F"]. */
export type ChannelOrder = string[];

/**
 * Read the selected trace format. InkML defaults to X and Y
 * when no channel order is declared.
 */
export function resolveChannelOrder(traceFormat: InkXmlObject | undefined): ChannelOrder {
	if (!traceFormat) {
		return ['X', 'Y'];
	}
	const channels = ensureArray(nsGet(traceFormat, 'channel'));
	const names = channels
		.map((channel) =>
			String(nsAttr(channel, 'name') ?? '')
				.trim()
				.toUpperCase(),
		)
		.filter((name) => name.length > 0);
	return names.length > 0 ? names : ['X', 'Y'];
}

/** Build an SVG path (`M x y L x y ...`) from decoded points and channel order. */
export function pointsToSvgPath(points: number[][], channelOrder: ChannelOrder): string {
	const xi = channelOrder.indexOf('X');
	const yi = channelOrder.indexOf('Y');
	const xIndex = xi >= 0 ? xi : 0;
	const yIndex = yi >= 0 ? yi : 1;
	const segments: string[] = [];
	for (const point of points) {
		const x = point[xIndex];
		const y = point[yIndex];
		if (!isFiniteNumber(x) || !isFiniteNumber(y)) {
			continue;
		}
		segments.push(`${segments.length === 0 ? 'M' : 'L'} ${x} ${y}`);
	}
	return segments.join(' ');
}

/** Extract normalised (0-1) pressure values from the F channel, when present. */
export function pointsToPressures(
	points: number[][],
	channelOrder: ChannelOrder,
	authored = false,
	range?: { min: number; max: number },
): number[] {
	const fIndex = channelOrder.indexOf('F');
	if (fIndex < 0) {
		return [];
	}
	const pressures: number[] = [];
	for (const point of points) {
		const raw = point[fIndex];
		if (isFiniteNumber(raw)) {
			// InkML force is commonly integer-encoded (0..32767); clamp to 0-1
			// for the renderer, normalising the typical integer range.
			// The declared range wins (a part may write 0..1023 or 0..8191).
			const normalised = range
				? (raw - range.min) / (range.max - range.min)
				: !authored && raw > 1
					? raw / 32767
					: raw;
			pressures.push(Math.max(0, Math.min(1, normalised)));
		}
	}
	return pressures;
}

/**
 * Extract per-point timestamps (milliseconds) from the trace's `T` channel,
 * when it declared one.
 *
 * This project's own real-world corpus has none: PowerPoint's own SaveAs
 * output declares only `X`/`Y` in its `traceFormat` (see the fixture behind
 * `contentpart-real-ink-roundtrip.test.ts`), and annotates the whole trace
 * GROUP with a single wall-clock `<inkml:timestamp>` rather than a per-point
 * channel. Other InkML producers (a captured digitizer session, OneNote, a
 * Surface Hub export) do author a genuine per-point `T` channel, and the
 * InkML spec's own default unit for it is milliseconds, so this is decoded
 * whenever present rather than assumed absent. Returns an empty array when no
 * `T` channel is declared, or no point carries a finite value, which is the
 * "no real timing data" signal `buildInkReplayTimeline`
 * (`pptx-viewer-shared`'s `render/ink-replay-timeline.ts`) falls back on.
 */
export function pointsToTimestamps(points: number[][], channelOrder: ChannelOrder): number[] {
	const tIndex = channelOrder.indexOf('T');
	if (tIndex < 0) {
		return [];
	}
	const timestamps: number[] = [];
	for (const point of points) {
		const raw = point[tIndex];
		if (isFiniteNumber(raw)) {
			timestamps.push(raw);
		}
	}
	return timestamps;
}

/** Decoded per-point pen-tilt lean, derived from whichever tilt channels a trace declared. */
export interface TiltChannels {
	/** Lean direction at each point, in radians (page-plane angle). */
	angles: number[];
	/** Lean strength at each point, 0 (upright) to 1 (maximally leaned). */
	magnitudes: number[];
	/**
	 * Which channel pair the source trace actually declared: the InkML-
	 * conventional `OTx`/`OTy` (or `tiltX`/`tiltY`) vector pair, or
	 * `AZIMUTH`/`ALTITUDE`. Recorded so a save that has to rewrite this part
	 * (see `inkml-content-part-writer.ts`) can re-declare the SAME channel
	 * pair instead of silently converting every re-saved stroke to `OTx`/`OTy`;
	 * the angle/magnitude values themselves are already a lossless decode of
	 * either encoding, so this field only affects which channel NAMES get
	 * written back, not the rendered lean.
	 */
	encoding: 'vector' | 'azimuthAltitude';
}

/**
 * Extract pen-tilt lean (direction + strength) from whichever tilt channels a
 * trace declares, or `undefined` when none of the recognised names are
 * present. Two encodings are recognised, in priority order:
 *
 * - A vector pair (`OTx`/`OTy`, the InkML-conventional tilt-offset names, or
 *   the `tiltX`/`tiltY` alias some digitizers use): the angle is the vector's
 *   direction, and the magnitude is its length, normalised against the
 *   largest magnitude seen anywhere in the trace (these are arbitrary device
 *   units with no fixed scale).
 * - `AZIMUTH` (degrees, the compass direction the pen leans toward), paired
 *   with an optional `ALTITUDE` (degrees from the writing surface: 90 is
 *   upright, 0 is flat). When altitude is absent, a fixed mid-range magnitude
 *   is used since azimuth alone gives a direction but not a lean strength.
 *
 * Every other declared channel (timing, or anything unrecognised) is still
 * decoded positionally by {@link decodeTracePoints} but has no consumer here,
 * matching the documented behaviour in `docs/guide/limitations.md`.
 */
export function pointsToTilt(
	points: number[][],
	channelOrder: ChannelOrder,
): TiltChannels | undefined {
	const otxIndex = firstIndexOf(channelOrder, ['OTX', 'TILTX']);
	const otyIndex = firstIndexOf(channelOrder, ['OTY', 'TILTY']);
	if (otxIndex >= 0 && otyIndex >= 0) {
		return tiltFromVector(points, otxIndex, otyIndex);
	}
	const azimuthIndex = channelOrder.indexOf('AZIMUTH');
	if (azimuthIndex >= 0) {
		return tiltFromAzimuthAltitude(points, azimuthIndex, channelOrder.indexOf('ALTITUDE'));
	}
	return undefined;
}

function firstIndexOf(channelOrder: ChannelOrder, names: string[]): number {
	for (const name of names) {
		const index = channelOrder.indexOf(name);
		if (index >= 0) {
			return index;
		}
	}
	return -1;
}

/** Derive angle + normalised magnitude from a per-point tilt-offset vector. */
function tiltFromVector(
	points: number[][],
	xIndex: number,
	yIndex: number,
): TiltChannels | undefined {
	const oxs: number[] = [];
	const oys: number[] = [];
	for (const point of points) {
		// A missing channel becomes NaN and is skipped by the finite check below.
		oxs.push(point[xIndex] ?? NaN);
		oys.push(point[yIndex] ?? NaN);
	}
	return tiltChannelsFromXY(oxs, oys);
}

/**
 * Derive per-point angle (radians) + magnitude (0-1, normalised against the
 * largest magnitude anywhere in `oxs`/`oys`) from parallel tilt-offset vector
 * arrays. Non-finite pairs are dropped (not zero-filled), matching
 * {@link tiltFromVector}'s original point-filtering behaviour.
 *
 * This is the shared core of {@link tiltFromVector} (InkML `OTx`/`OTy`
 * channel decode) and of the InkML writer's inverse mapping in
 * `inkml-content-part.ts` (`tiltOffsetAt`): both need the exact same
 * vector<->angle/magnitude convention so a written value reads back to the
 * same angle/magnitude. It is also mirrored (small, dependency-free, no
 * shared code to import across the core/shared package boundary) as
 * `tiltChannelsFromVectors` in `pptx-viewer-shared`'s `render/ink-tilt-nib.ts`
 * for the Draw tool's own live-captured `tiltX`/`tiltY` degrees.
 */
export function tiltChannelsFromXY(
	oxs: readonly number[],
	oys: readonly number[],
): TiltChannels | undefined {
	const angles: number[] = [];
	const rawMagnitudes: number[] = [];
	const len = Math.min(oxs.length, oys.length);
	for (let i = 0; i < len; i++) {
		const ox = oxs[i];
		const oy = oys[i];
		if (!isFiniteNumber(ox) || !isFiniteNumber(oy)) {
			continue;
		}
		angles.push(Math.atan2(oy, ox));
		rawMagnitudes.push(Math.hypot(ox, oy));
	}
	if (angles.length === 0) {
		return undefined;
	}
	const max = Math.max(0, ...rawMagnitudes);
	const magnitudes = rawMagnitudes.map((m) => (max > 0 ? Math.min(1, m / max) : 0));
	return { angles, magnitudes, encoding: 'vector' };
}

/** Derive angle + magnitude from AZIMUTH (required) and ALTITUDE (optional), both in degrees. */
function tiltFromAzimuthAltitude(
	points: number[][],
	azimuthIndex: number,
	altitudeIndex: number,
): TiltChannels | undefined {
	const angles: number[] = [];
	const magnitudes: number[] = [];
	for (const point of points) {
		const azimuth = point[azimuthIndex];
		if (!isFiniteNumber(azimuth)) {
			continue;
		}
		angles.push((azimuth * Math.PI) / 180);
		const altitude = altitudeIndex >= 0 ? point[altitudeIndex] : undefined;
		magnitudes.push(
			isFiniteNumber(altitude) ? Math.max(0, Math.min(1, 1 - (altitude as number) / 90)) : 0.5,
		);
	}
	if (angles.length === 0) {
		return undefined;
	}
	return { angles, magnitudes, encoding: 'azimuthAltitude' };
}

/** Read a child value by local element name, ignoring any XML namespace prefix. */
export function nsGet(obj: InkXmlObject, localName: string): unknown {
	if (localName in obj) {
		return obj[localName];
	}
	for (const key of Object.keys(obj)) {
		if (localNameOf(key) === localName && !key.startsWith('@_')) {
			return obj[key];
		}
	}
	return undefined;
}

/** Read an attribute value by local name, ignoring the `@_` prefix and any ns. */
export function nsAttr(obj: InkXmlObject, localName: string): unknown {
	const direct = obj[`@_${localName}`];
	if (direct !== undefined) {
		return direct;
	}
	for (const key of Object.keys(obj)) {
		if (key.startsWith('@_') && localNameOf(key.slice(2)) === localName) {
			return obj[key];
		}
	}
	return undefined;
}

/** Strip a leading `prefix:` namespace qualifier from an element/attribute key. */
export function localNameOf(key: string): string {
	const colon = key.indexOf(':');
	return colon >= 0 ? key.slice(colon + 1) : key;
}

/** Coerce a possibly-single XML node (or absent value) into an array. */
export function ensureArray(value: unknown): InkXmlObject[] {
	if (value === undefined || value === null) {
		return [];
	}
	return (Array.isArray(value) ? value : [value]) as InkXmlObject[];
}
