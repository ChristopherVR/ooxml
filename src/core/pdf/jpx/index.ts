/**
 * A JPEG 2000 Part 1 decoder (ITU-T T.800 | ISO/IEC 15444-1) for the PDF reader's JPXDecode
 * images: JP2/JPX files or bare codestreams, every progression order, tiles and tile-parts,
 * packed packet headers, all code-block styles, reversible and irreversible paths and the
 * component transforms. High Throughput (Part 15) codestreams are refused.
 */
import { readJp2, type JpxHeader } from './boxes';
import { decodeCodestream, type JpxComponent } from './decode';
import { DEFAULT_LIMITS, JpxError, type JpxLimits } from './errors';

export type { JpxChannel, JpxHeader, JpxPalette } from './boxes';
export type { JpxComponent } from './decode';
export { JpxError, type JpxLimits } from './errors';

export interface JpxImage {
	/** The image area on the reference grid (its origin may be offset). */
	x0: number;
	y0: number;
	width: number;
	height: number;
	components: JpxComponent[];
	/** The JP2 header boxes (empty for a bare codestream). */
	header: JpxHeader;
}

/** Decodes JPEG 2000 data; throws a `JpxError` saying why it cannot. */
export function decodeJpx(data: Uint8Array, limits: Partial<JpxLimits> = {}): JpxImage {
	try {
		const { codestream, header } = readJp2(data);
		const planes = decodeCodestream(codestream, { ...DEFAULT_LIMITS, ...limits });
		return { ...planes, header };
	} catch (error) {
		if (error instanceof JpxError) throw error;
		if (error instanceof RangeError && /allocation|length/i.test(error.message))
			throw new JpxError('it is larger than the memory available');
		throw new JpxError('its JPEG 2000 data is damaged');
	}
}
