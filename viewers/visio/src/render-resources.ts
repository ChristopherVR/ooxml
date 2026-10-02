import type { VisioImage } from 'ooxml-core/visio';
/** One URL per shared core image resource, with explicit per-render disposal. */
export class RenderResources {
	#images = new Map<Uint8Array, Map<string, string>>();
	#disposed = false;
	imageUrl(image: VisioImage): string {
		if (this.#disposed) throw new Error('Render resources have been disposed.');
		let types = this.#images.get(image.bytes);
		if (!types) {
			types = new Map();
			this.#images.set(image.bytes, types);
		}
		let url = types.get(image.mimeType);
		if (!url) {
			url = URL.createObjectURL(new Blob([Uint8Array.from(image.bytes)], { type: image.mimeType }));
			types.set(image.mimeType, url);
		}
		return url;
	}
	dispose(): void {
		if (this.#disposed) return;
		this.#disposed = true;
		for (const types of this.#images.values())
			for (const url of types.values()) URL.revokeObjectURL(url);
		this.#images.clear();
	}
}
