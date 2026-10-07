import { afterEach, expect, it, vi } from 'vitest';
import { demoDocument } from 'ooxml-core/visio/ui';
import { renderPage } from './render-svg';
import { rasterFixture } from './__fixtures__/raster-fixtures.mjs';
import { createRotationPreview } from './viewer-rotation-preview';
afterEach(() => {
	document.body.replaceChildren();
	vi.unstubAllGlobals();
});
it('uses the live renderer without source/history changes and restores the selected source', () => {
	const model = structuredClone(demoDocument);
	const page = model.pages[0]!;
	const shape = page.shapes[0]!;
	shape.rotation = { pinX: 3, pinY: 5, angle: 0 };
	const before = structuredClone(model);
	const rendered = renderPage(model, page);
	document.body.append(rendered.svg);
	const source = rendered.svg.querySelector<SVGGElement>('[data-shape-id]')!;
	const originalTransform = source.getAttribute('transform');
	const preview = createRotationPreview(model, page, shape, source);
	preview.update(Math.PI / 2);
	expect(source.style.visibility).toBe('hidden');
	expect(source.getAttribute('transform')).toBe(originalTransform);
	expect(preview.element.querySelectorAll('[data-shape-id]')).toHaveLength(0);
	expect(preview.element.querySelector('path')!.getAttribute('d')).toBe(
		source.querySelector('path')!.getAttribute('d'),
	);
	expect(preview.element.querySelector('g')!.getAttribute('transform')).not.toBe(originalTransform);
	preview.update(Math.PI / 3);
	expect(rendered.svg.querySelectorAll('.rotation-shape-preview')).toHaveLength(1);
	expect(model).toEqual(before);
	preview.dispose();
	expect(source.style.visibility).toBe('');
	expect(rendered.svg.querySelectorAll('.rotation-shape-preview')).toHaveLength(0);
	rendered.dispose();
});

it('releases replaced and cancelled temporary raster URLs while retaining the source URL', () => {
	const createObjectURL = vi.fn(() => `blob:rotation-${createObjectURL.mock.calls.length}`);
	const revokeObjectURL = vi.fn();
	vi.stubGlobal('URL', Object.assign(class extends URL {}, { createObjectURL, revokeObjectURL }));
	const model = structuredClone(demoDocument);
	const page = model.pages[0]!;
	const shape = page.shapes[0]!;
	shape.rotation = { pinX: 3, pinY: 5, angle: 0 };
	shape.image = {
		...rasterFixture(),
		x: 0,
		y: 0,
		width: 1,
		height: 1,
	};
	const rendered = renderPage(model, page);
	document.body.append(rendered.svg);
	const source = rendered.svg.querySelector<SVGGElement>('[data-shape-id]')!;
	const preview = createRotationPreview(model, page, shape, source);
	preview.update(Math.PI / 6);
	preview.update(Math.PI / 3);
	expect(createObjectURL).toHaveBeenCalledTimes(3);
	expect(revokeObjectURL.mock.calls).toEqual([['blob:rotation-2']]);
	preview.dispose();
	expect(revokeObjectURL.mock.calls).toEqual([['blob:rotation-2'], ['blob:rotation-3']]);
	rendered.dispose();
	expect(revokeObjectURL.mock.calls).toEqual([
		['blob:rotation-2'],
		['blob:rotation-3'],
		['blob:rotation-1'],
	]);
});
