/**
 * Decodes a JPEG 2000 codestream tile by tile: packets (tier 2), code-blocks (tier 1), the
 * inverse wavelet transform, the inverse multiple component transform (RCT or ICT, G.2 and G.3)
 * and the DC level shift (G.1), into one plane of samples per component.
 */
import { readCodestream, type Codestream, type TileData } from './codestream';
import { inverseDwt } from './dwt';
import { Budget, JpxError, type JpxLimits } from './errors';
import { HeaderBits, readPacket, Truncated, type PacketSource } from './packets';
import { progress } from './progression';
import { buildTile, type TileComponent } from './structure';
import { decodeBlock } from './tier1';

export interface JpxComponent {
	/** Size of the component on its own (sub-sampled) grid. */
	width: number;
	height: number;
	dx: number;
	dy: number;
	depth: number;
	signed: boolean;
	/** Samples in their own range (0..2^depth - 1, or signed). */
	data: Int32Array;
}

export interface JpxPlanes {
	/** The image area on the reference grid. */
	x0: number;
	y0: number;
	width: number;
	height: number;
	components: JpxComponent[];
}

const concat = (parts: Uint8Array[]): Uint8Array => {
	if (parts.length === 1) return parts[0]!;
	const out = new Uint8Array(parts.reduce((sum, part) => sum + part.length, 0));
	let at = 0;
	for (const part of parts) {
		out.set(part, at);
		at += part.length;
	}
	return out;
};

/** The samples of one tile-component, before the component transform and level shift. */
function componentSamples(component: TileComponent, budget: Budget): Float32Array {
	const coefficients = component.resolutions.map((res) =>
		res.bands.map((band) => {
			const width = band.x1 - band.x0;
			const out = new Float32Array(Math.max(0, width * (band.y1 - band.y0)));
			for (const block of band.blocks)
				decodeBlock(
					block,
					band,
					component.style.style,
					component.roiShift,
					component.style.reversible,
					out,
					width,
					budget,
				);
			return out;
		}),
	);
	return inverseDwt(component, coefficients, budget);
}

function inverseMct(samples: Float32Array[], reversible: boolean): void {
	const [a, b, c] = samples as [Float32Array, Float32Array, Float32Array];
	if (a.length !== b.length || a.length !== c.length) return;
	for (let i = 0; i < a.length; i++) {
		const y = a[i]!;
		const u = b[i]!;
		const v = c[i]!;
		if (reversible) {
			const g = y - Math.floor((u + v) / 4);
			a[i] = v + g;
			b[i] = g;
			c[i] = u + g;
		} else {
			a[i] = y + 1.402 * v;
			b[i] = y - 0.34413 * u - 0.71414 * v;
			c[i] = y + 1.772 * u;
		}
	}
}

function decodeTile(
	stream: Codestream,
	data: TileData,
	out: JpxComponent[],
	origins: { x: number; y: number }[],
	budget: Budget,
): void {
	const tile = buildTile(stream, data.index, data.markers);
	const source: PacketSource = { data: concat(data.parts), at: 0 };
	if (data.headers) source.headers = new HeaderBits(concat(data.headers), 0);
	try {
		progress(tile, budget, (layer, c, r, k) => readPacket(tile, layer, c, r, k, source, budget));
	} catch (error) {
		if (!(error instanceof Truncated)) throw error;
	}
	const samples = tile.components.map((component) => componentSamples(component, budget));
	if (tile.style.mct === 1 && samples.length >= 3)
		inverseMct(samples, tile.components[0]!.style.reversible);
	tile.components.forEach((component, c) => {
		const target = out[c]!;
		const width = component.x1 - component.x0;
		const shift = target.signed ? 0 : 2 ** (target.depth - 1);
		const min = target.signed ? -(2 ** (target.depth - 1)) : 0;
		const max = target.signed ? 2 ** (target.depth - 1) - 1 : 2 ** target.depth - 1;
		const values = samples[c]!;
		for (let y = component.y0; y < component.y1; y++) {
			const row = (y - origins[c]!.y) * target.width - origins[c]!.x;
			for (let x = component.x0; x < component.x1; x++) {
				const value = Math.round(values[(y - component.y0) * width + x - component.x0]! + shift);
				target.data[row + x] = value < min ? min : value > max ? max : value;
			}
		}
	});
}

export function decodeCodestream(data: Uint8Array, limits: JpxLimits): JpxPlanes {
	const stream = readCodestream(data, limits);
	const { siz } = stream;
	const budget = new Budget(limits.maxWork);
	if (siz.components.some((info) => info.depth > 30))
		throw new JpxError('its JPEG 2000 sample size is not supported');
	const origins = siz.components.map((info) => ({
		x: Math.ceil(siz.x0 / info.dx),
		y: Math.ceil(siz.y0 / info.dy),
	}));
	const components = siz.components.map((info, c): JpxComponent => {
		const width = Math.ceil(siz.x1 / info.dx) - origins[c]!.x;
		const height = Math.ceil(siz.y1 / info.dy) - origins[c]!.y;
		return { width, height, ...info, data: new Int32Array(width * height) };
	});
	for (const tile of stream.tiles) decodeTile(stream, tile, components, origins, budget);
	return { x0: siz.x0, y0: siz.y0, width: siz.x1 - siz.x0, height: siz.y1 - siz.y0, components };
}
