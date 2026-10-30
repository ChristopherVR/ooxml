import { zlibSync } from 'fflate';

/**
 * The zlib (RFC 1950) stream of `data`, deflated synchronously. PDF's `FlateDecode` filter is this
 * format; `fflate` is the deflate the core already ships, so no second compressor is needed.
 */
export function zlibCompress(data: Uint8Array): Uint8Array {
	return zlibSync(data);
}
