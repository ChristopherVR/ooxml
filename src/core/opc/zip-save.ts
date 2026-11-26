import type JSZip from 'jszip';

// Container, image and font formats that already compress their payloads. Keep formats such as
// SVG, TIFF, WAV, TTF and obfuscated OpenType on the usual ZIP compression path.
const COMPRESSED_MEDIA = /\.(?:png|jpe?g|gif|webp|mp[34]|m4[av]|aac|ogg|opus|webm|woff2?)$/iu;

/** Keep loaded DEFLATE streams and use STORE for already compressed media added to a package. */
export function keepCompressedMedia(zip: JSZip): void {
	for (const entry of Object.values(zip.files)) {
		if (entry.dir || entry.options.compression || !COMPRESSED_MEDIA.test(entry.name)) continue;
		// JSZip exposes the requested compression setting on an entry, but leaves it unset on load.
		// Its CompressedObject keeps the original ZIP method as two little-endian bytes.
		const data = '_data' in entry ? entry._data : undefined;
		if (typeof data === 'object' && data && 'compression' in data) {
			const method = data.compression;
			if (
				typeof method !== 'object' ||
				!method ||
				!('magic' in method) ||
				method.magic !== '\0\0' ||
				!('compressedSize' in data) ||
				!('uncompressedSize' in data) ||
				data.compressedSize !== data.uncompressedSize
			)
				continue;
		}
		entry.options.compression = 'STORE';
	}
}

/** Keep the original entry when serializing it would emit exactly the same UTF-8 text. */
export function writeZipText(
	zip: JSZip,
	path: string,
	text: string,
	previousText: string | undefined,
): void {
	// A decoder may produce U+FFFD from malformed bytes. Equality after a lossy decode does not
	// prove equality with the bytes JSZip would write.
	if (text !== previousText || text.includes('\uFFFD')) zip.file(path, text);
}

/** Keep an opaque source part if the model still has the same bytes. */
export async function writeZipBytes(zip: JSZip, path: string, bytes: Uint8Array): Promise<void> {
	const existing = await zip.file(path)?.async('uint8array');
	// A stored entry can expose a view into the archive. A caller may have edited that
	// shared buffer in place, so JSZip must recalculate the entry's checksum.
	if (
		existing &&
		existing.buffer !== bytes.buffer &&
		existing.length === bytes.length &&
		existing.every((value, i) => value === bytes[i])
	)
		return;
	zip.file(path, bytes);
}
