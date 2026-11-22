const GLB_MAGIC = 0x46546c67;
const JSON_CHUNK = 0x4e4f534a;

/** Recognise a GLB 2.0 whose textures and buffers do not depend on other files. */
export function isSelfContainedGlb(bytes: Uint8Array): boolean {
	if (bytes.byteLength < 24) return false;
	const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
	if (
		view.getUint32(0, true) !== GLB_MAGIC ||
		view.getUint32(4, true) !== 2 ||
		view.getUint32(8, true) !== bytes.byteLength ||
		view.getUint32(16, true) !== JSON_CHUNK
	)
		return false;

	const jsonSize = view.getUint32(12, true);
	if (jsonSize % 4 || jsonSize > bytes.byteLength - 20) return false;
	for (let offset = 20 + jsonSize; offset < bytes.byteLength;) {
		if (offset + 8 > bytes.byteLength) return false;
		const size = view.getUint32(offset, true);
		if (size % 4 || size > bytes.byteLength - offset - 8) return false;
		offset += 8 + size;
	}
	try {
		const manifest: unknown = JSON.parse(
			new TextDecoder('utf-8', { fatal: true }).decode(bytes.subarray(20, 20 + jsonSize)),
		);
		if (!manifest || typeof manifest !== 'object' || !('asset' in manifest)) return false;
		const asset = manifest.asset;
		if (!asset || typeof asset !== 'object' || !('version' in asset) || asset.version !== '2.0') {
			return false;
		}
		// Check nested extensions too, including URI keys written as JSON escapes.
		const nodes: unknown[] = [manifest];
		while (nodes.length > 0) {
			const node = nodes.pop();
			if (!node || typeof node !== 'object') continue;
			if (Object.hasOwn(node, 'uri')) return false;
			for (const value of Object.values(node)) {
				if (value && typeof value === 'object') nodes.push(value);
			}
		}
		return true;
	} catch {
		return false;
	}
}
