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
		return { has, readText, icon: (path: string) => pngPreview(files.get(path)) };
	} catch (error) {
		if (error instanceof AddInManifestError) throw error;
		if (error instanceof ZipLimitError && error.kind !== 'partName')
			return fail('LIMIT_INFLATE', 'archive entry');
		return fail('INVALID_ZIP', 'archive');
	}
}

function pngPreview(bytes: Uint8Array | undefined): string | undefined {
	if (!bytes || bytes.length < 33 || bytes.length > MANIFEST_BYTES) return;
	const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
	if (
		view.getUint32(0) !== 0x89504e47 ||
		view.getUint32(4) !== 0x0d0a1a0a ||
		view.getUint32(8) !== 13 ||
		view.getUint32(12) !== 0x49484452
	)
		return;
	const width = view.getUint32(16);
	const height = view.getUint32(20);
	if (!width || !height || width > 512 || height > 512) return;
	let binary = '';
	for (let offset = 0; offset < bytes.length; offset += 8192)
		binary += String.fromCharCode(...bytes.subarray(offset, offset + 8192));
	return `data:image/png;base64,${btoa(binary)}`;
}
