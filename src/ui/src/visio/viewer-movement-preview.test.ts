import { afterEach, expect, it, vi } from 'vitest';
import { demoDocument, visioMovePreviewTransform } from 'ooxml-core/visio/ui';
import { parseVsdx } from 'ooxml-core/visio';
import JSZip from 'jszip';
import { createVsdxFixture } from './__fixtures__/fixture.mjs';
import { renderPage } from './render-svg';
import { createMovementPreview } from './viewer-movement-preview';

afterEach(() => {
	document.body.replaceChildren();
	vi.restoreAllMocks();
});
it('shares inert pattern and gradient definitions exactly like a translated redraw, and cleans up cancellation', async () => {
	const model = structuredClone(demoDocument),
		page = model.pages[0]!;
	page.shapes = page.shapes.filter((shape) => shape.kind === 'shape').slice(0, 2);
	const shape = page.shapes[0]!;
	const zip = await JSZip.loadAsync(await createVsdxFixture());
	const xml = await zip.file('visio/pages/page1.xml')!.async('string');
	zip.file(
		'visio/pages/page1.xml',
		xml.replace('<Text>', '<Cell N="FillPattern" V="2"/><Cell N="FillBkgnd" V="#ffffff"/><Text>'),
	);
	shape.style.fillPattern = (
		await parseVsdx(await zip.generateAsync({ type: 'uint8array' }))
	).pages[0]!.shapes[0]!.style.fillPattern!;
	page.shapes[1]!.style.fillGradient = {
		type: 'linear',
		start: [0, 0],
		end: [1, 1],
		stops: [
			{ offset: 0, color: '#fff', opacity: 1 },
			{ offset: 1, color: '#f00', opacity: 1 },
		],
	};
	const url = vi.spyOn(URL, 'createObjectURL').mockReturnValue('blob:pattern');
	const revoke = vi.spyOn(URL, 'revokeObjectURL').mockImplementation(() => {});
	const original = renderPage(model, page);
	document.body.append(original.svg);
	const groups = Array.from(original.svg.querySelectorAll<SVGGElement>('[data-shape-id]'));
	const before = original.svg.querySelector('defs')!.innerHTML;
	const preview = createMovementPreview(
		original.svg,
		page,
		page.shapes.map((shape) => shape.id),
	);
	const delta = { x: 1.25, y: -0.37 };
	preview.update(delta);
	const clones = Array.from(original.svg.querySelectorAll<SVGGElement>('[data-movement-preview]'));
	for (const [index, clone] of clones.entries()) {
		expect(clone.querySelector('path')!.getAttribute('fill')).toBe(
			groups[index]!.querySelector('path')!.getAttribute('fill'),
		);
		expect(clone.querySelectorAll('[data-shape-id], [tabindex], [role]')).toHaveLength(0);
	}
	expect(original.svg.querySelector('defs')!.innerHTML).toBe(before);
	expect(url).toHaveBeenCalledTimes(1);
	const movedPage = {
		...page,
		shapes: page.shapes.map((shape) => ({
			...shape,
			transform: visioMovePreviewTransform(shape, delta)!,
		})),
	};
	const redraw = renderPage({ ...model, pages: [movedPage] }, movedPage);
	expect(redraw.svg.querySelector('pattern')!.getAttribute('patternTransform')).toBe(
		original.svg.querySelector('pattern')!.getAttribute('patternTransform'),
	);
	expect(redraw.svg.querySelector('linearGradient')!.innerHTML).toBe(
		original.svg.querySelector('linearGradient')!.innerHTML,
	);
	expect(clones[0]!.getAttribute('transform')).toBe(
		redraw.svg.querySelector('[data-shape-id]')!.getAttribute('transform'),
	);
	preview.dispose();
	expect(original.svg.querySelector('[data-movement-preview]')).toBeNull();
	expect(groups.every((group) => group.style.visibility === '')).toBe(true);
	expect(original.svg.querySelector('defs')!.innerHTML).toBe(before);
	expect(revoke).not.toHaveBeenCalled();
	original.dispose();
	redraw.dispose();
	expect(revoke).toHaveBeenCalledTimes(2);
});
