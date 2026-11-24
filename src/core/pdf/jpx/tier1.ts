/**
 * Tier-1 decoding (ITU-T T.800 Annex D): the significance propagation, magnitude refinement and
 * cleanup passes over a code-block's stripes, with the code-block style options: arithmetic
 * coding bypass, context reset, termination on every pass, vertically causal contexts and
 * segmentation symbols. The decoded indices are dequantized (E.1) into the band's coefficients,
 * reconstructed at the middle of the undecoded range, after the ROI max-shift is undone.
 */
import { initialContexts, RUN_LENGTH, UNIFORM, ZERO_CODING } from './contexts';
import type { Budget } from './errors';
import { MqDecoder, RawDecoder } from './mq';
import type { Band, CodeBlock, Segment } from './structure';

const SIG = 1;
const VISIT = 2;
const REFINED = 4;
const NEG = 8;

const flags = new Uint8Array(6400);
const magnitudes = new Float64Array(4096);
const lows = new Int8Array(4096);
const contexts = new Uint8Array(19);

const BYPASS = 1;
const RESET = 2;
const VCAUSAL = 8;
const SEGSYM = 32;

function joined(segment: Segment): Uint8Array {
	if (segment.chunks.length === 1) return segment.chunks[0]!;
	const out = new Uint8Array(segment.chunks.reduce((sum, chunk) => sum + chunk.length, 0));
	let at = 0;
	for (const chunk of segment.chunks) {
		out.set(chunk, at);
		at += chunk.length;
	}
	return out;
}

/** Decodes a code-block into `out` (the band's coefficients, `stride` wide). */
export function decodeBlock(
	block: CodeBlock,
	band: Band,
	style: number,
	roiShift: number,
	reversible: boolean,
	out: Float32Array,
	stride: number,
	budget: Budget,
): void {
	const w = block.x1 - block.x0;
	const h = block.y1 - block.y0;
	const planes = band.magnitudeBits + roiShift - block.zeroPlanes;
	const total = Math.min(block.passes, 3 * planes - 2);
	if (w <= 0 || h <= 0 || planes <= 0 || total <= 0 || !block.segments.length) return;
	const s = w + 2;
	flags.fill(0, 0, s * (h + 2));
	magnitudes.fill(0, 0, w * h);
	initialContexts(contexts);
	const zc = ZERO_CODING[band.orientation]!;
	const vcausal = (style & VCAUSAL) !== 0;
	let mq: MqDecoder | undefined;
	let raw: RawDecoder | undefined;
	let segmentIndex = -1;
	let left = 0;

	// Significant neighbours of padded index p, as a zero coding table index.
	const neighbours = (p: number, south: boolean): number => {
		const hz = (flags[p - 1]! & SIG) + (flags[p + 1]! & SIG);
		let v = flags[p - s]! & SIG;
		let d = (flags[p - s - 1]! & SIG) + (flags[p - s + 1]! & SIG);
		if (south) {
			v += flags[p + s]! & SIG;
			d += (flags[p + s - 1]! & SIG) + (flags[p + s + 1]! & SIG);
		}
		return hz * 15 + v * 5 + d;
	};
	const contribution = (f: number) => (f & SIG ? (f & NEG ? -1 : 1) : 0);
	const sign = (p: number, south: boolean): number => {
		if (raw) return raw.bit();
		const hz = Math.max(-1, Math.min(1, contribution(flags[p - 1]!) + contribution(flags[p + 1]!)));
		const v = Math.max(
			-1,
			Math.min(1, contribution(flags[p - s]!) + (south ? contribution(flags[p + s]!) : 0)),
		);
		const label = 3 * hz + v;
		return label >= 0 ? mq!.decode(contexts, 9 + label) : mq!.decode(contexts, 9 - label) ^ 1;
	};
	const bit = (label: number) => (raw ? raw.bit() : mq!.decode(contexts, label));
	const significant = (p: number, i: number, south: boolean, plane: number) => {
		flags[p]! |= SIG | (sign(p, south) ? NEG : 0);
		magnitudes[i] = 2 ** plane;
		lows[i] = plane;
	};

	for (let pass = 0; pass < total; pass++) {
		if (left === 0) {
			const segment = block.segments[++segmentIndex];
			if (!segment) break;
			left = segment.passes;
			const bypassed = (style & BYPASS) !== 0 && pass >= 10 && pass % 3 !== 0;
			raw = bypassed ? new RawDecoder(joined(segment)) : undefined;
			mq = bypassed ? undefined : new MqDecoder(joined(segment));
		}
		left--;
		budget.spend(w * h);
		const kind = pass === 0 ? 0 : pass % 3; // 0 cleanup, 1 significance, 2 refinement
		const plane = planes - 1 - Math.floor((pass + 2) / 3);
		for (let y0 = 0; y0 < h; y0 += 4) {
			const y1 = Math.min(y0 + 4, h);
			for (let x = 0; x < w; x++) {
				let y = y0;
				if (kind === 0 && y1 - y0 === 4) {
					let quiet = true;
					for (let row = y0; row < y1 && quiet; row++) {
						const p = (row + 1) * s + x + 1;
						quiet = !flags[p] && !neighbours(p, !vcausal || row !== y1 - 1);
					}
					if (quiet) {
						if (!mq!.decode(contexts, RUN_LENGTH)) continue;
						y += (mq!.decode(contexts, UNIFORM) << 1) | mq!.decode(contexts, UNIFORM);
						significant((y + 1) * s + x + 1, y * w + x, !vcausal || (y & 3) !== 3, plane);
						y++;
					}
				}
				for (; y < y1; y++) {
					const p = (y + 1) * s + x + 1;
					const i = y * w + x;
					const f = flags[p]!;
					const south = !vcausal || (y & 3) !== 3;
					if (kind === 2) {
						if ((f & (SIG | VISIT)) !== SIG) continue;
						const label = f & REFINED ? 16 : neighbours(p, south) ? 15 : 14;
						if (bit(label)) magnitudes[i]! += 2 ** plane;
						lows[i] = plane;
						flags[p] = f | REFINED;
						continue;
					}
					if (f & (SIG | VISIT)) continue;
					const n = neighbours(p, south);
					if (kind === 1 && !n) continue;
					if (bit(zc[n]!)) significant(p, i, south, plane);
					if (kind === 1) flags[p]! |= VISIT;
				}
			}
		}
		if (kind === 0) {
			if (style & SEGSYM) for (let k = 0; k < 4; k++) mq!.decode(contexts, UNIFORM);
			for (let p = 0; p < s * (h + 2); p++) flags[p]! &= ~VISIT;
		}
		if (style & RESET) initialContexts(contexts);
	}
	const shift = 2 ** roiShift;
	for (let y = 0; y < h; y++)
		for (let x = 0; x < w; x++) {
			const i = y * w + x;
			let m = magnitudes[i]!;
			if (!m) continue;
			let low = lows[i]!;
			if (roiShift && m >= shift) {
				m = Math.floor(m / shift);
				low = Math.max(0, low - roiShift);
			}
			const value = (m + (reversible ? (low > 0 ? 2 ** (low - 1) : 0) : 2 ** low / 2)) * band.delta;
			const at = (block.y0 - band.y0 + y) * stride + block.x0 - band.x0 + x;
			out[at] = flags[(y + 1) * s + x + 1]! & NEG ? -value : value;
		}
}
