import { parseLegacyVsd } from './legacy.js';
import { parseVsdx, type ParseVsdxOptions } from './parser.js';
import { VisioPackageError } from './package.js';
import type { VisioDocument } from './model.js';

/** Dispatch on container bytes, keeping the VSDX parser and its options intact. */
export async function loadVisio(
	input: Uint8Array | ArrayBuffer,
	options: ParseVsdxOptions = {},
): Promise<VisioDocument> {
	const bytes = input instanceof Uint8Array ? input : new Uint8Array(input);
	if (
		[0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1].every((value, index) => bytes[index] === value)
	)
		return parseLegacyVsd(bytes, options);
	if (bytes[0] === 0x50 && bytes[1] === 0x4b) return parseVsdx(input, options);
	throw new VisioPackageError(
		'UNSUPPORTED_FORMAT',
		'Expected a VSD compound file or VSDX ZIP package.',
	);
}
