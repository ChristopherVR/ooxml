import { crc32 } from '../../digest/crc32';
import { decodeZipPath, inspectZipHeaders } from '../../opc/zip-headers';
import { openZip, ZipLimitError } from '../../opc/zip-limits';
import { AddInManifestError, MANIFEST_BYTES, PACKAGE_BYTES } from './validation';

const limits = {
	maxInputBytes: PACKAGE_BYTES,
	maxEntries: 128,
	maxEntryBytes: PACKAGE_BYTES,
	maxTotalBytes: PACKAGE_BYTES,
	maxCompressionRatio: 100,
};
const utf8 = new TextDecoder('utf-8', { fatal: true });

export function decodeManifest(bytes: Uint8Array): string {
	if (bytes.byteLength > MANIFEST_BYTES) throw new AddInManifestError('limit', 'manifest size');
	try {
		return utf8.decode(bytes);
	} catch {
		throw new AddInManifestError('manifest', 'UTF-8');
	}
}

export async function openManifestPackage(bytes: Uint8Array) {
	const fail = (code: string, field: string): never => {
		throw new AddInManifestError(code.startsWith('LIMIT_') ? 'limit' : 'package', field);
	};
	try {
		const headers = inspectZipHeaders(bytes, limits, (code) => fail(code, code));
		const zip = await openZip(bytes, limits);
		const names = Object.keys(zip.files);
		if (names.length !== headers.size || names.some((name) => !headers.has(name)))
			fail('ZIP_MISMATCH', 'archive index');
		const paths = new Set<string>();
		for (const name of names) {
			const key = decodeZipPath(name, fail).replace(/\/$/, '').normalize('NFC').toLowerCase();
			if (name.length > 240 || paths.has(key)) fail('UNSAFE_PATH', 'archive path');
			paths.add(key);
		}
		const has = (path: string) => {
			const entry = headers.get(path);
			return entry !== undefined && !entry.directory && entry.size > 0;
		};
		const files = new Map<string, Uint8Array>();
		// Check the actual output and CRC of every entry before accepting the package. No asset is decoded.
		for (const [name, header] of headers) {
			if (header.directory) continue;
			const file = zip.file(name);
			if (!file) return fail('ZIP_MISMATCH', 'archive entry');
			const data = await file.async('uint8array');
			if (data.byteLength !== header.size || (crc32(data) ^ 0xffffffff) >>> 0 !== header.crc)
				fail('ZIP_MISMATCH', 'archive checksum');
			files.set(name, data);
		}
		const readText = async (path: string) => {
			const data = files.get(path);
			if (!has(path) || !data) return fail('MISSING_ENTRY', 'manifest or locale file');
			return decodeManifest(data);
		};
		return {
			has,
			readText,
			async icon(path: string) {
				const bytes = files.get(path);
				if (!bytes || bytes[0] !== 0x89 || bytes.length > MANIFEST_BYTES) return;
				const { pngPreview } = await import('./package-icon');
				return pngPreview(bytes);
			},
		};
	} catch (error) {
		if (error instanceof AddInManifestError) throw error;
		if (error instanceof ZipLimitError && error.kind !== 'partName')
			return fail('LIMIT_INFLATE', 'archive entry');
		return fail('INVALID_ZIP', 'archive');
	}
}
