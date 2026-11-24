/**
 * Region encoders for the JBIG2 test encoder: generic regions (templates 0 to 3, TPGDON) and
 * generic refinement regions (templates 0 and 1, TPGRON), with context templates written out from
 * ITU-T T.88 6.2.5.3 and 6.3.5.3 independently of the decoder's tables.
 */
import { MqEncoder } from './jbig2-encode';

type At = [number, number];

/** Context of pixel (x, y) for a generic template (6.2.5.3), spelled out per template. */
function genericContext(
	p: (x: number, y: number) => number,
	x: number,
	y: number,
	t: number,
	at: At[],
): number {
	const a = (i: number) => p(x + at[i]![0], y + at[i]![1]);
	const bits: number[] =
		t === 0
			? [
					a(3),
					p(x - 1, y - 2),
					p(x, y - 2),
					p(x + 1, y - 2),
					a(2),
					a(1),
					p(x - 2, y - 1),
					p(x - 1, y - 1),
					p(x, y - 1),
					p(x + 1, y - 1),
					p(x + 2, y - 1),
					a(0),
					p(x - 4, y),
					p(x - 3, y),
					p(x - 2, y),
					p(x - 1, y),
				]
			: t === 1
				? [
						p(x - 1, y - 2),
						p(x, y - 2),
						p(x + 1, y - 2),
						p(x + 2, y - 2),
						p(x - 2, y - 1),
						p(x - 1, y - 1),
						p(x, y - 1),
						p(x + 1, y - 1),
						p(x + 2, y - 1),
						a(0),
						p(x - 3, y),
						p(x - 2, y),
						p(x - 1, y),
					]
				: t === 2
					? [
							p(x - 1, y - 2),
							p(x, y - 2),
							p(x + 1, y - 2),
							p(x - 2, y - 1),
							p(x - 1, y - 1),
							p(x, y - 1),
							p(x + 1, y - 1),
							a(0),
							p(x - 2, y),
							p(x - 1, y),
						]
					: [
							p(x - 3, y - 1),
							p(x - 2, y - 1),
							p(x - 1, y - 1),
							p(x, y - 1),
							p(x + 1, y - 1),
							a(0),
							p(x - 4, y),
							p(x - 3, y),
							p(x - 2, y),
							p(x - 1, y),
						];
	return bits.reduce((context, bit) => (context << 1) | bit, 0);
}

/** Arithmetic-codes `rows` (strings of '0' and '1', 1 black) as a generic region. */
export function encodeGeneric(
	rows: string[],
	template: number,
	at: At[],
	tpgdon: boolean,
): Uint8Array {
	const encoder = new MqEncoder();
	encodeGenericInto(encoder, new Uint8Array(1 << 16), rows, template, at, tpgdon);
	return encoder.finish();
}

/** Codes a generic region with an encoder and contexts that continue (symbol dictionaries). */
export function encodeGenericInto(
	encoder: MqEncoder,
	contexts: Uint8Array,
	rows: string[],
	template: number,
	at: At[],
	tpgdon: boolean,
): void {
	const width = rows[0]?.length ?? 0;
	const p = (x: number, y: number) =>
		x < 0 || y < 0 || x >= width || y >= rows.length ? 0 : rows[y]![x] === '1' ? 1 : 0;
	const sltp = [0x9b25, 0x0795, 0x00e5, 0x0195][template]!;
	let ltp = 0;
	for (let y = 0; y < rows.length; y++) {
		if (tpgdon) {
			const same = rows[y] === (y ? rows[y - 1] : '0'.repeat(width)) ? 1 : 0;
			encoder.encode(contexts, sltp, same ^ ltp);
			ltp = same;
			if (same) continue;
		}
		for (let x = 0; x < width; x++)
			encoder.encode(contexts, genericContext(p, x, y, template, at), p(x, y));
	}
}

/**
 * Arithmetic-codes `rows` as a generic refinement region of `reference` (6.3.5, template 0 with
 * the given AT pixels or template 1, reference offset 0), with TPGRON when asked.
 */
export function encodeRefinement(
	rows: string[],
	reference: string[],
	template: number,
	at: At[],
	tpgron: boolean,
): Uint8Array {
	const width = rows[0]?.length ?? 0;
	const pick = (source: string[]) => (x: number, y: number) =>
		x < 0 || y < 0 || y >= source.length || x >= (source[y]?.length ?? 0)
			? 0
			: source[y]![x] === '1'
				? 1
				: 0;
	const g = pick(rows);
	const r = pick(reference);
	const context = (x: number, y: number): number => {
		const bits =
			template === 0
				? [
						r(x + at[1]![0], y + at[1]![1]),
						r(x, y - 1),
						r(x + 1, y - 1),
						r(x - 1, y),
						r(x, y),
						r(x + 1, y),
						r(x - 1, y + 1),
						r(x, y + 1),
						r(x + 1, y + 1),
						g(x + at[0]![0], y + at[0]![1]),
						g(x, y - 1),
						g(x + 1, y - 1),
						g(x - 1, y),
					]
				: [
						r(x, y - 1),
						r(x - 1, y),
						r(x, y),
						r(x + 1, y),
						r(x, y + 1),
						r(x + 1, y + 1),
						g(x - 1, y - 1),
						g(x, y - 1),
						g(x + 1, y - 1),
						g(x - 1, y),
					];
		return bits.reduce((value, bit) => (value << 1) | bit, 0);
	};
	const uniform = (x: number, y: number) => {
		for (let j = -1; j <= 1; j++)
			for (let i = -1; i <= 1; i++) if (r(x + i, y + j) !== r(x, y)) return false;
		return true;
	};
	const contexts = new Uint8Array(1 << 13);
	const encoder = new MqEncoder();
	let ltp = 0;
	for (let y = 0; y < rows.length; y++) {
		if (tpgron) {
			// A row is typical when every uniform-neighbourhood pixel equals its reference pixel.
			let typical = 1;
			for (let x = 0; x < width; x++) if (uniform(x, y) && g(x, y) !== r(x, y)) typical = 0;
			encoder.encode(contexts, template ? 1 << 7 : 1 << 8, typical ^ ltp);
			ltp = typical;
		}
		for (let x = 0; x < width; x++) {
			if (ltp && uniform(x, y)) continue;
			encoder.encode(contexts, context(x, y), g(x, y));
		}
	}
	return encoder.finish();
}
