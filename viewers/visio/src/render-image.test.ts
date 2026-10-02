import { describe, it, expect, vi, afterEach } from 'vitest';
import { demoDocument } from './demo-document.js';
import { renderPage } from './render-svg.js';
import { mountViewer } from './binding.js';

function imageDocument() {
	const model = structuredClone(demoDocument);
	model.pages[0]!.shapes[0]!.image = {
		mimeType: 'image/png',
		bytes: new Uint8Array([137, 80, 78, 71]),
		pixelWidth: 1,
		pixelHeight: 1,
		x: 0.1,
		y: 0.2,
		width: 0.5,
		height: 0.7,
	};
	return model;
}
afterEach(() => {
	vi.unstubAllGlobals();
	document.body.replaceChildren();
});
describe('raster resource lifecycle', () => {
	function mockUrls() {
		const create = vi.fn(() => 'blob:local-safe-image'),
			revoke = vi.fn();
		vi.stubGlobal(
			'URL',
			Object.assign(class extends URL {}, { createObjectURL: create, revokeObjectURL: revoke }),
		);
		return { create, revoke };
	}
	it('creates bounded local image URL and disposes it exactly once', () => {
		const { create, revoke } = mockUrls();
		const model = imageDocument();
		const result = renderPage(model, model.pages[0]!);
		expect(create).toHaveBeenCalledTimes(1);
		expect(result.svg.querySelector('image')?.getAttribute('href')).toBe('blob:local-safe-image');
		expect(result.svg.querySelector('image')?.getAttribute('transform')).toBe(
			'translate(0.1 0.8999999999999999) scale(1 -1)',
		);
		result.dispose();
		result.dispose();
		expect(revoke).toHaveBeenCalledTimes(1);
	});
	it('revokes replaced/disconnected images and recreates URLs on reattach', () => {
		const { create, revoke } = mockUrls();
		const host = document.createElement('div');
		document.body.append(host);
		const viewer = mountViewer(host, { document: imageDocument() });
		expect(create).toHaveBeenCalledTimes(1);
		viewer.element.remove();
		expect(revoke).toHaveBeenCalledTimes(1);
		host.append(viewer.element);
		expect(create).toHaveBeenCalledTimes(2);
		viewer.update({ document: demoDocument });
		expect(revoke).toHaveBeenCalledTimes(2);
		viewer.destroy();
	});
});

describe('shared image decode resources', () => {
	it('shares one object URL for repeated image bytes and revokes it once', () => {
		const create = vi.fn(() => 'blob:shared'),
			revoke = vi.fn();
		vi.stubGlobal(
			'URL',
			Object.assign(class extends URL {}, { createObjectURL: create, revokeObjectURL: revoke }),
		);
		const model = imageDocument();
		const original = model.pages[0]!.shapes[0]!;
		const copy = { ...original, id: 'shared-image' };
		model.pages[0]!.shapes.push(copy);
		const result = renderPage(model, model.pages[0]!);
		expect(create).toHaveBeenCalledOnce();
		expect(result.svg.querySelectorAll('image')).toHaveLength(2);
		result.dispose();
		expect(revoke).toHaveBeenCalledOnce();
	});
});

describe('suspended presentation', () => {
	it('does not recreate image resources for updates while disconnected', () => {
		const create = vi.fn(() => 'blob:deferred'),
			revoke = vi.fn();
		vi.stubGlobal(
			'URL',
			Object.assign(class extends URL {}, { createObjectURL: create, revokeObjectURL: revoke }),
		);
		const host = document.createElement('div');
		document.body.append(host);
		const viewer = mountViewer(host, { document: imageDocument() });
		viewer.element.remove();
		viewer.update({ document: imageDocument(), zoom: 2 });
		expect(create).toHaveBeenCalledOnce();
		expect(revoke).toHaveBeenCalledOnce();
		host.append(viewer.element);
		expect(create).toHaveBeenCalledTimes(2);
		expect(viewer.element.zoom).toBe(2);
		viewer.destroy();
	});
});
