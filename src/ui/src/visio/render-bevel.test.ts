import { afterEach, expect, it } from 'vitest';
import { setupFormattingViewer as setup } from './__fixtures__/formatting-viewer';
import { renderPage } from './render-svg';

afterEach(() => document.body.replaceChildren());

const table = (svg: Element) =>
	svg
		.querySelector('feComponentTransfer feFuncA')!
		.getAttribute('tableValues')!
		.split(' ')
		.map(Number);

it('lights a height field built from the bevel type, the lighting and the material', async () => {
	const ui = await setup();
	ui.selection();
	ui.press('bevel-angle');
	await ui.done();
	const document = ui.controller.state.document!;
	const angle = renderPage(document, document.pages[0]!);
	const filter = angle.svg.querySelector('filter')!;
	expect([...filter.children].map((node) => node.localName).slice(0, 3)).toEqual([
		'feGaussianBlur',
		'feComponentTransfer',
		'feDiffuseLighting',
	]);
	const heights = table(angle.svg);
	expect(heights).toHaveLength(128);
	expect(Math.max(...heights)).toBe(1);
	const light = filter.querySelector('feDistantLight')!;
	expect(Number.isFinite(Number(light.getAttribute('azimuth')))).toBe(true);
	// The group no longer fades the overlay: the filter's alphas carry the strength.
	expect(angle.svg.querySelector('[data-bevel]')!.hasAttribute('opacity')).toBe(false);
	expect(angle.svg.querySelector('[data-bevel]')!.getAttribute('data-bevel-type')).toBe('5');
	angle.dispose();

	ui.press('bevel-circle');
	await ui.done();
	const circle = renderPage(ui.controller.state.document!, ui.controller.state.document!.pages[0]!);
	expect(table(circle.svg)).not.toEqual(heights);
	circle.dispose();

	// A saved lighting type moves the light; the flat material draws no bevel at all.
	const shape = ui.controller.state.document!.pages[0]!.shapes[0]!;
	const first = filter.querySelector('feDistantLight')!.getAttribute('azimuth');
	shape.style.bevel = { ...shape.style.bevel!, lighting: 5 };
	const relit = renderPage(ui.controller.state.document!, ui.controller.state.document!.pages[0]!);
	expect(relit.svg.querySelector('feDistantLight')!.getAttribute('azimuth')).not.toBe(first);
	relit.dispose();
	shape.style.bevel = { ...shape.style.bevel!, material: 7 };
	const flat = renderPage(ui.controller.state.document!, ui.controller.state.document!.pages[0]!);
	expect(flat.svg.querySelector('[data-bevel]')).toBeNull();
	flat.dispose();
	shape.style.bevel = { ...shape.style.bevel!, material: 10 };
	const powder = renderPage(ui.controller.state.document!, ui.controller.state.document!.pages[0]!);
	expect(powder.svg.querySelector('filter feFlood')).not.toBeNull();
	powder.dispose();
	ui.dispose();
	ui.controller.destroy();
});

it('fits the pinhole projection to the text of a shape drawn from projected paths', async () => {
	const ui = await setup();
	ui.selection();
	ui.press('bevel-circle');
	await ui.done();
	ui.press('rotation-3d-perspective-left');
	await ui.done();
	const document = ui.controller.state.document!;
	const perspective = renderPage(document, document.pages[0]!);
	const text = perspective.svg.querySelector('[data-solid-flat]')!;
	const numbers = (text.getAttribute('transform') ?? '')
		.replace(/^matrix\(|\)$/g, '')
		.split(' ')
		.map(Number);
	expect(numbers).toHaveLength(6);
	expect(numbers.every(Number.isFinite)).toBe(true);
	expect(text.querySelector('text, tspan, [data-text]')).not.toBeNull();
	perspective.dispose();
	// The same text under the parallel isometric preset takes the parallel matrix instead.
	ui.press('rotation-3d-isometric-right-up');
	await ui.done();
	const parallel = renderPage(
		ui.controller.state.document!,
		ui.controller.state.document!.pages[0]!,
	);
	expect(parallel.svg.querySelector('[data-rotation3d]')!.getAttribute('transform')).toMatch(
		/^matrix\(/,
	);
	parallel.dispose();
	ui.dispose();
	ui.controller.destroy();
});

it('turns the children of a group with it', async () => {
	const ui = await setup();
	const document = ui.controller.state.document!;
	const group = document.pages[0]!.shapes[0]!;
	const child = structuredClone(group);
	child.id = `${group.id}-child`;
	child.children = [];
	group.kind = 'group';
	group.children = [child];
	group.style = {
		...group.style,
		rotation3d: { x: 0.5, y: 0.2, z: 0, perspective: Math.PI / 3, keepTextFlat: false },
	};
	const rendered = renderPage(document, document.pages[0]!);
	const host = rendered.svg.querySelector('[data-group3d]')!;
	expect(host.hasAttribute('data-group3d-perspective')).toBe(true);
	expect(host.getAttribute('transform')).toMatch(/^matrix\(/);
	// The child's own drawing sits inside the turned container.
	expect(host.querySelector('path')).not.toBeNull();
	rendered.dispose();
	// A group that is not turned has no container.
	group.style = { ...group.style };
	delete (group.style as { rotation3d?: unknown }).rotation3d;
	const plain = renderPage(document, document.pages[0]!);
	expect(plain.svg.querySelector('[data-group3d]')).toBeNull();
	plain.dispose();
	ui.dispose();
	ui.controller.destroy();
});
