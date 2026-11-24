/**
 * The geometry of a tile (ITU-T T.800 B.2 to B.7): tile-components, resolution levels, sub-bands,
 * precincts and code-blocks, with the quantization of each sub-band (E.1). Precincts are built
 * when their first packet is read, so a sparse codestream costs only what it codes.
 */
import type { Codestream } from './codestream';
import { JpxError } from './errors';
import type {
	ComponentStyle,
	Markers,
	ProgressionChange,
	Quantization,
	TileStyle,
} from './markers';
import { TagTree } from './tagtree';

export interface Segment {
	chunks: Uint8Array[];
	length: number;
	passes: number;
	maxPasses: number;
}

export interface CodeBlock {
	x0: number;
	y0: number;
	x1: number;
	y1: number;
	included: boolean;
	lblock: number;
	zeroPlanes: number;
	passes: number;
	segments: Segment[];
}

export interface Band {
	/** 0 LL, 1 HL, 2 LH, 3 HH. */
	orientation: number;
	x0: number;
	y0: number;
	x1: number;
	y1: number;
	/** Magnitude bit-planes Mb (E-2). */
	magnitudeBits: number;
	/** Quantization step (1 when reversible). */
	delta: number;
	/** Code-block size exponents in this resolution. */
	xcb: number;
	ycb: number;
	blocks: CodeBlock[];
}

export interface PrecinctBand {
	band: Band;
	blocksWide: number;
	blocks: CodeBlock[];
	inclusion: TagTree;
	zeroPlanes: TagTree;
}

export interface Resolution {
	x0: number;
	y0: number;
	x1: number;
	y1: number;
	ppx: number;
	ppy: number;
	precinctsWide: number;
	precinctsHigh: number;
	bands: Band[];
	precincts: Map<number, PrecinctBand[]>;
}

export interface TileComponent {
	x0: number;
	y0: number;
	x1: number;
	y1: number;
	/** Sub-sampling on the reference grid. */
	dx: number;
	dy: number;
	style: ComponentStyle;
	roiShift: number;
	resolutions: Resolution[];
}

export interface Tile {
	index: number;
	x0: number;
	y0: number;
	x1: number;
	y1: number;
	style: TileStyle;
	components: TileComponent[];
	progression: ProgressionChange[];
}

const ceilDiv = (a: number, b: number) => Math.ceil(a / b);

function bandQuantization(
	quant: Quantization,
	depth: number,
	r: number,
	orientation: number,
	reversible: boolean,
): { magnitudeBits: number; delta: number } {
	const index = r === 0 ? 0 : 1 + 3 * (r - 1) + (orientation - 1);
	let exponent: number;
	let mantissa: number;
	if (quant.style === 1) {
		exponent = quant.exponents[0]! - (r === 0 ? 0 : r - 1);
		mantissa = quant.mantissas[0]!;
	} else {
		if (index >= quant.exponents.length)
			throw new JpxError('its JPEG 2000 quantization is incomplete');
		exponent = quant.exponents[index]!;
		mantissa = quant.mantissas[index]!;
	}
	const gain = orientation === 0 ? 0 : orientation === 3 ? 2 : 1;
	const delta = reversible ? 1 : 2 ** (depth + gain - exponent) * (1 + mantissa / 2048);
	return { magnitudeBits: quant.guard + exponent - 1, delta };
}

/** The tile's markers over the main header's (B.2's precedence for COD, COC, QCD and QCC). */
function pick<T>(
	component: number,
	tile: Markers,
	main: Markers,
	perComponent: (m: Markers) => Map<number, T>,
	whole: (m: Markers) => T | undefined,
): T {
	const value =
		perComponent(tile).get(component) ??
		whole(tile) ??
		perComponent(main).get(component) ??
		whole(main);
	return value!;
}

export function buildTile(stream: Codestream, index: number, markers: Markers): Tile {
	const { siz, main } = stream;
	const p = index % siz.tilesWide;
	const q = Math.floor(index / siz.tilesWide);
	const x0 = Math.max(siz.tileX0 + p * siz.tileWidth, siz.x0);
	const y0 = Math.max(siz.tileY0 + q * siz.tileHeight, siz.y0);
	const x1 = Math.min(siz.tileX0 + (p + 1) * siz.tileWidth, siz.x1);
	const y1 = Math.min(siz.tileY0 + (q + 1) * siz.tileHeight, siz.y1);
	const components = siz.components.map((info, c): TileComponent => {
		const style = pick(
			c,
			markers,
			main,
			(m) => m.componentStyles,
			(m) => m.style,
		);
		const quant = pick(
			c,
			markers,
			main,
			(m) => m.componentQuantization,
			(m) => m.quantization,
		);
		const roiShift = markers.roiShift.get(c) ?? main.roiShift.get(c) ?? 0;
		const tc = {
			x0: ceilDiv(x0, info.dx),
			y0: ceilDiv(y0, info.dy),
			x1: ceilDiv(x1, info.dx),
			y1: ceilDiv(y1, info.dy),
		};
		const levels = style.levels;
		const resolutions = Array.from({ length: levels + 1 }, (_, r): Resolution => {
			const scale = 2 ** (levels - r);
			const res = {
				x0: ceilDiv(tc.x0, scale),
				y0: ceilDiv(tc.y0, scale),
				x1: ceilDiv(tc.x1, scale),
				y1: ceilDiv(tc.y1, scale),
			};
			const ppx = style.ppx[r]!;
			const ppy = style.ppy[r]!;
			if (r > 0 && (!ppx || !ppy)) throw new JpxError('its JPEG 2000 precinct size is invalid');
			const xcb = Math.min(style.xcb, r ? ppx - 1 : ppx);
			const ycb = Math.min(style.ycb, r ? ppy - 1 : ppy);
			const bands = (r === 0 ? [0] : [1, 2, 3]).map((orientation): Band => {
				const nb = r === 0 ? levels : levels - r + 1;
				const xob = orientation & 1;
				const yob = orientation >> 1;
				const half = r === 0 ? 0 : 2 ** (nb - 1);
				return {
					orientation,
					x0: ceilDiv(tc.x0 - half * xob, 2 ** nb),
					y0: ceilDiv(tc.y0 - half * yob, 2 ** nb),
					x1: ceilDiv(tc.x1 - half * xob, 2 ** nb),
					y1: ceilDiv(tc.y1 - half * yob, 2 ** nb),
					xcb,
					ycb,
					blocks: [],
					...bandQuantization(quant, info.depth, r, orientation, style.reversible),
				};
			});
			const empty = res.x1 <= res.x0 || res.y1 <= res.y0;
			return {
				...res,
				ppx,
				ppy,
				precinctsWide: empty ? 0 : ceilDiv(res.x1, 2 ** ppx) - Math.floor(res.x0 / 2 ** ppx),
				precinctsHigh: empty ? 0 : ceilDiv(res.y1, 2 ** ppy) - Math.floor(res.y0 / 2 ** ppy),
				bands,
				precincts: new Map(),
			};
		});
		return { ...tc, dx: info.dx, dy: info.dy, style, roiShift, resolutions };
	});
	const style = markers.tileStyle ?? main.tileStyle!;
	const progression = markers.progression.length ? markers.progression : main.progression;
	return { index, x0, y0, x1, y1, style, components, progression };
}

/** The bands of precinct `k` of a resolution, with their code-blocks (built on first use). */
export function precinct(resolution: Resolution, k: number): PrecinctBand[] {
	const known = resolution.precincts.get(k);
	if (known) return known;
	const px = k % resolution.precinctsWide;
	const py = Math.floor(k / resolution.precinctsWide);
	const width = 2 ** resolution.ppx;
	const height = 2 ** resolution.ppy;
	const rx0 = (Math.floor(resolution.x0 / width) + px) * width;
	const ry0 = (Math.floor(resolution.y0 / height) + py) * height;
	const scale = resolution.bands[0]!.orientation === 0 ? 1 : 2;
	const bands = resolution.bands.map((band): PrecinctBand => {
		const bx0 = Math.max(band.x0, rx0 / scale);
		const by0 = Math.max(band.y0, ry0 / scale);
		const bx1 = Math.min(band.x1, (rx0 + width) / scale);
		const by1 = Math.min(band.y1, (ry0 + height) / scale);
		const blocks: CodeBlock[] = [];
		let blocksWide = 0;
		let blocksHigh = 0;
		if (bx1 > bx0 && by1 > by0) {
			const cw = 2 ** band.xcb;
			const ch = 2 ** band.ycb;
			const cx0 = Math.floor(bx0 / cw);
			const cy0 = Math.floor(by0 / ch);
			blocksWide = ceilDiv(bx1, cw) - cx0;
			blocksHigh = ceilDiv(by1, ch) - cy0;
			for (let y = 0; y < blocksHigh; y++)
				for (let x = 0; x < blocksWide; x++) {
					const block: CodeBlock = {
						x0: Math.max(bx0, (cx0 + x) * cw),
						y0: Math.max(by0, (cy0 + y) * ch),
						x1: Math.min(bx1, (cx0 + x + 1) * cw),
						y1: Math.min(by1, (cy0 + y + 1) * ch),
						included: false,
						lblock: 3,
						zeroPlanes: 0,
						passes: 0,
						segments: [],
					};
					blocks.push(block);
					band.blocks.push(block);
				}
		}
		return {
			band,
			blocksWide,
			blocks,
			inclusion: new TagTree(blocksWide, blocksHigh),
			zeroPlanes: new TagTree(blocksWide, blocksHigh),
		};
	});
	resolution.precincts.set(k, bands);
	return bands;
}
