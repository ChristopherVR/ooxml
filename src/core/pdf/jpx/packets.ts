/**
 * Packets (ITU-T T.800 B.9 and B.10): the packet header bits with their bit stuffing, code-block
 * inclusion and zero bit-planes through tag trees, the number of coding passes, Lblock and the
 * codeword segment lengths, then the code-block data of the packet body. SOP and EPH markers are
 * skipped, and headers are read from PPM or PPT data when the codestream packs them.
 */
import type { Budget } from './errors';
import { precinct, type CodeBlock, type Segment, type Tile } from './structure';

/** Thrown when packet data ends early; the tile keeps what was read before. */
export class Truncated extends Error {}

export class HeaderBits {
	private last = 0;
	private count = 0;

	constructor(
		readonly data: Uint8Array,
		public at: number,
	) {}

	bit(): number {
		if (this.count === 0) {
			if (this.at >= this.data.length) throw new Truncated();
			this.count = this.last === 0xff ? 7 : 8;
			this.last = this.data[this.at++]!;
		}
		this.count--;
		return (this.last >> this.count) & 1;
	}

	bits(n: number): number {
		let value = 0;
		for (let i = 0; i < n; i++) value = value * 2 + this.bit();
		return value;
	}

	/** Ends a header: skips the rest of the byte and the stuffed byte after a 0xFF. */
	align(): void {
		if (this.last === 0xff) this.at++;
		this.last = 0;
		this.count = 0;
	}

	/** Skips a marker (EPH) at the read position. */
	skipMarker(code: number): void {
		if (this.data[this.at] === 0xff && this.data[this.at + 1] === (code & 0xff)) this.at += 2;
	}
}

export interface PacketSource {
	data: Uint8Array;
	at: number;
	/** Packed packet headers (PPM, PPT). */
	headers?: HeaderBits;
}

function passCount(header: HeaderBits): number {
	if (!header.bit()) return 1;
	if (!header.bit()) return 2;
	let value = header.bits(2);
	if (value < 3) return 3 + value;
	value = header.bits(5);
	if (value < 31) return 6 + value;
	return 37 + header.bits(7);
}

/** The next codeword segment of a block for new passes (D.4.1: termall and bypass split them). */
function openSegment(block: CodeBlock, style: number): Segment {
	const last = block.segments[block.segments.length - 1];
	if (last && last.passes < last.maxPasses) return last;
	let maxPasses = Infinity;
	if (style & 4) maxPasses = 1;
	else if (style & 1)
		maxPasses = !last ? 10 : last.maxPasses === 1 || last.maxPasses === 10 ? 2 : 1;
	const segment = { chunks: [], length: 0, passes: 0, maxPasses };
	block.segments.push(segment);
	return segment;
}

/** Reads the packet of layer `layer` for precinct `k` of resolution `r` of component `c`. */
export function readPacket(
	tile: Tile,
	layer: number,
	c: number,
	r: number,
	k: number,
	source: PacketSource,
	budget: Budget,
): void {
	const component = tile.components[c]!;
	const resolution = component.resolutions[r]!;
	const { data } = source;
	if (tile.style.sop && data[source.at] === 0xff && data[source.at + 1] === 0x91) source.at += 6;
	const header = source.headers ?? new HeaderBits(data, source.at);
	const bands = precinct(resolution, k);
	budget.spend(1 + bands.reduce((sum, band) => sum + band.blocks.length, 0));
	const contributions: { block: CodeBlock; passes: number; parts: [Segment, number, number][] }[] =
		[];
	if (header.bit()) {
		const style = component.style.style;
		for (const pb of bands)
			pb.blocks.forEach((block, index) => {
				const x = index % pb.blocksWide;
				const y = Math.floor(index / pb.blocksWide);
				const bit = () => header.bit();
				const included = block.included
					? header.bit() === 1
					: pb.inclusion.decode(x, y, layer + 1, bit) !== undefined;
				if (!included) return;
				if (!block.included) {
					let threshold = 1;
					while (pb.zeroPlanes.decode(x, y, threshold, bit) === undefined) {
						if (++threshold > 80) throw new Truncated();
					}
					block.zeroPlanes = threshold - 1;
					block.included = true;
				}
				const passes = passCount(header);
				while (header.bit()) if (++block.lblock > 32) throw new Truncated();
				const parts: [Segment, number, number][] = [];
				for (let left = passes; left > 0;) {
					const segment = openSegment(block, style);
					const count = Math.min(segment.maxPasses - segment.passes, left);
					const length = header.bits(block.lblock + Math.floor(Math.log2(count)));
					segment.passes += count;
					parts.push([segment, count, length]);
					left -= count;
				}
				contributions.push({ block, passes, parts });
			});
	}
	header.align();
	if (tile.style.eph) header.skipMarker(0xff92);
	if (!source.headers) source.at = header.at;
	for (const { block, passes, parts } of contributions) {
		block.passes += passes;
		for (const [segment, , length] of parts) {
			const end = source.at + length;
			segment.chunks.push(data.subarray(source.at, Math.min(end, data.length)));
			segment.length += length;
			source.at = end;
			if (end > data.length) throw new Truncated();
		}
	}
}
