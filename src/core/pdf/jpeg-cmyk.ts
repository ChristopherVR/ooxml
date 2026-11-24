/**
 * CMYK JPEGs as a document can hold them. Applications that open a JPEG on its own read CMYK
 * samples as inverted when the file has an Adobe APP14 segment and as they are otherwise; a PDF
 * viewer does the same and then applies the image's /Decode array. The picture keeps the PDF's
 * colours by adding or removing the APP14 segment, so no pixel is decoded or re-encoded.
 */

const SOI = 0xd8;
const APP14 = 0xee;
const SOS = 0xda;

interface Segment {
	marker: number;
	start: number;
	end: number;
}

function segments(bytes: Uint8Array): Segment[] {
	const out: Segment[] = [];
	if (bytes[0] !== 0xff || bytes[1] !== SOI) return out;
	for (let at = 2; at + 4 <= bytes.length && out.length < 1000;) {
		if (bytes[at] !== 0xff) break;
		const marker = bytes[at + 1]!;
		if (marker === 0xff) {
			at++;
			continue;
		}
		const length = (bytes[at + 2]! << 8) | bytes[at + 3]!;
		if (length < 2) break;
		out.push({ marker, start: at, end: at + 2 + length });
		if (marker === SOS) break;
		at += 2 + length;
	}
	return out;
}

/** The colour transform of the Adobe APP14 segment (0 none, 1 YCbCr, 2 YCCK), if there is one. */
export function adobeTransform(bytes: Uint8Array): number | undefined {
	for (const segment of segments(bytes)) {
		if (segment.marker !== APP14 || segment.end - segment.start < 16) continue;
		const at = segment.start + 4;
		const tag = String.fromCharCode(...bytes.subarray(at, at + 5));
		if (tag === 'Adobe') return bytes[segment.start + 15];
	}
	return undefined;
}

const ADOBE_SEGMENT = Uint8Array.of(
	0xff,
	APP14,
	0x00,
	0x0e,
	0x41,
	0x64,
	0x6f,
	0x62,
	0x65,
	0x00,
	0x64,
	0x00,
	0x00,
	0x00,
	0x00,
	0x00,
);

/**
 * The CMYK JPEG to store so that it shows as the PDF shows it, given whether its /Decode array
 * inverts the components; undefined when that needs a decode (YCCK data with an inverting array).
 */
export function cmykJpegForDocument(
	bytes: Uint8Array,
	decodeInverted: boolean,
): Uint8Array | undefined {
	if (!decodeInverted) return bytes;
	const transform = adobeTransform(bytes);
	if (transform === undefined) {
		const out = new Uint8Array(bytes.length + ADOBE_SEGMENT.length);
		out.set(bytes.subarray(0, 2));
		out.set(ADOBE_SEGMENT, 2);
		out.set(bytes.subarray(2), 2 + ADOBE_SEGMENT.length);
		return out;
	}
	if (transform !== 0) return undefined;
	const segment = segments(bytes).find(
		(item) =>
			item.marker === APP14 &&
			String.fromCharCode(...bytes.subarray(item.start + 4, item.start + 9)) === 'Adobe',
	)!;
	const out = new Uint8Array(bytes.length - (segment.end - segment.start));
	out.set(bytes.subarray(0, segment.start));
	out.set(bytes.subarray(segment.end), segment.start);
	return out;
}
