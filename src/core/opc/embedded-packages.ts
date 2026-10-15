import { openZip, type ZipLimits } from './zip-limits';

/** Limits for a host package and each embedded package opened inside it. */
const EMBEDDED_ZIP_LIMITS: ZipLimits = {
	label: 'Embedded',
	maxEntries: 20_000,
	maxInputBytes: 200 * 1024 * 1024,
	maxEntryBytes: 200 * 1024 * 1024,
	maxTotalBytes: 500 * 1024 * 1024,
};

/** Limits for the nested package itself (default 100 MiB). */
const NESTED_ZIP_LIMITS: ZipLimits = { label: 'Embedded' };

export interface EmbeddedPackage {
	path: string;
	name: string;
	kind: 'docx' | 'xlsx' | 'pptx' | 'vsdx' | 'unsupported';
	bytes: Uint8Array;
}

/** Direct package parts only. CFB-wrapped OLE objects are deliberately not treated as ZIP files. */
export async function listEmbeddedPackages(bytes: Uint8Array): Promise<EmbeddedPackage[]> {
	const zip = await openZip(bytes, EMBEDDED_ZIP_LIMITS);
	const result: EmbeddedPackage[] = [];
	for (const file of Object.values(zip.files)) {
		if (file.dir || !/^(word|ppt|xl|visio)\/embeddings\/[^/]+$/u.test(file.name)) continue;
		const ext = file.name.split('.').pop()?.toLowerCase();
		const kind =
			ext === 'docx' || ext === 'xlsx' || ext === 'pptx' || ext === 'vsdx' ? ext : 'unsupported';
		result.push({
			path: file.name,
			name: file.name.split('/').pop()!,
			kind,
			bytes: await file.async('uint8array'),
		});
	}
	return result;
}

/** Replace an existing native embedded package without rebuilding its parent document model. */
export async function replaceEmbeddedPackage(
	parent: Uint8Array,
	path: string,
	child: Uint8Array,
): Promise<Uint8Array> {
	if (!/^(word|ppt|xl|visio)\/embeddings\/[^/]+\.(docx|xlsx|pptx|vsdx)$/u.test(path))
		throw new Error('Only native Office package embeddings can be saved back');
	const zip = await openZip(parent, EMBEDDED_ZIP_LIMITS);
	if (!zip.file(path)) throw new Error('The embedded document no longer exists');
	const nested = await openZip(child, NESTED_ZIP_LIMITS);
	if (!nested.file('[Content_Types].xml'))
		throw new Error('The replacement is not an Office package');
	zip.file(path, child);
	return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
}
