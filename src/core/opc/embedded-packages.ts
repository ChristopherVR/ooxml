import JSZip from 'jszip';

export interface EmbeddedPackage {
	path: string;
	name: string;
	kind: 'docx' | 'xlsx' | 'pptx' | 'vsdx' | 'unsupported';
	bytes: Uint8Array;
}

/** Direct package parts only. CFB-wrapped OLE objects are deliberately not treated as ZIP files. */
export async function listEmbeddedPackages(bytes: Uint8Array): Promise<EmbeddedPackage[]> {
	const zip = await JSZip.loadAsync(bytes);
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
	const zip = await JSZip.loadAsync(parent);
	if (!zip.file(path)) throw new Error('The embedded document no longer exists');
	const nested = await JSZip.loadAsync(child);
	if (!nested.file('[Content_Types].xml'))
		throw new Error('The replacement is not an Office package');
	zip.file(path, child);
	return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
}
