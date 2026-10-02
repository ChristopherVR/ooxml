import type { VisioImage } from 'ooxml-core/visio';
import { svgElement } from './render-svg.js';

const MAX_EMBEDDED_BYTES = 16 * 1024 * 1024;
/** One reference per shared core image resource, with explicit per-render disposal. */
export class RenderResources {
	#images = new Map<Uint8Array, Map<string, string>>();
	#disposed = false;
	#embeddedBytes = 0;
	#imageId = 0;
	constructor(private readonly definitions?: SVGDefsElement) {}
	get portable(): boolean {
		return !!this.definitions;
	}
	imageUrl(image: VisioImage): string {
		if (this.#disposed) throw new Error('Render resources have been disposed.');
		let types = this.#images.get(image.bytes);
		if (!types) {
			types = new Map();
			this.#images.set(image.bytes, types);
		}
		let url = types.get(image.mimeType);
		if (!url) {
			if (this.definitions) {
				// Check before base64 allocation. Shared byte arrays are encoded only once.
				const size = 4 * Math.ceil(image.bytes.byteLength / 3);
				if (this.#embeddedBytes + size > MAX_EMBEDDED_BYTES)
					throw new Error('Embedded raster resources exceed the safe SVG export limit.');
				this.#embeddedBytes += size;
				const symbol = svgElement('symbol');
				symbol.id = `visio-export-image-${++this.#imageId}`;
				symbol.setAttribute('viewBox', '0 0 1 1');
				symbol.setAttribute('preserveAspectRatio', 'none');
				const node = svgElement('image');
				node.setAttribute('width', '1');
				node.setAttribute('height', '1');
				node.setAttribute('preserveAspectRatio', 'none');
				node.setAttributeNS(
					'http://www.w3.org/1999/xlink',
					'xlink:href',
					`data:${image.mimeType};base64,${base64(image.bytes)}`,
				);
				symbol.append(node);
				this.definitions.append(symbol);
				url = `#${symbol.id}`;
			} else
				url = URL.createObjectURL(
					new Blob([Uint8Array.from(image.bytes)], { type: image.mimeType }),
				);
			types.set(image.mimeType, url);
		}
		return url;
	}
	dispose(): void {
		if (this.#disposed) return;
		this.#disposed = true;
		if (!this.portable)
			for (const types of this.#images.values())
				for (const url of types.values()) URL.revokeObjectURL(url);
		this.#images.clear();
	}
}

function base64(bytes: Uint8Array): string {
	const chunks: string[] = [];
	// A multiple of three keeps all but the last chunk free of base64 padding.
	for (let offset = 0; offset < bytes.length; offset += 24_576)
		chunks.push(btoa(String.fromCharCode(...bytes.subarray(offset, offset + 24_576))));
	return chunks.join('');
}
