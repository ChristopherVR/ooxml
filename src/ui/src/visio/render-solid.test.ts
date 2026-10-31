import { afterEach, expect, it } from 'vitest';
import { setupFormattingViewer as setup } from './__fixtures__/formatting-viewer';
import { renderPage } from './render-svg';
import { shadeColor } from './render-solid';

afterEach(() => document.body.replaceChildren());

it('shades a colour below and above its own value', () => {
	expect(shadeColor('#ff8040', 0.5)).toBe('#804020');
	expect(shadeColor('#000000', 1.5)).toBe('#808080');
	expect(shadeColor('#336699', 1)).toBe('#336699');
});

it('projects perspective, draws the extrusion and contour, and keeps the outline hit-testable', async () => {
	const ui = await setup();
	ui.selection();
	ui.press('bevel-circle');
	await ui.done();
	ui.press('rotation-3d-perspective-left');
	await ui.done();
	ui.commands.run({
		type: 'shape-format',
		patch: {
			extrusionDepth: 20,
			extrusionColor: '#ff0000',
			contourSize: 4,
			contourColor: '#008000',
		},
	});
	await ui.done();
	expect(ui.shape().style.extrusion?.color).toBe('#ff0000');
	const document = ui.controller.state.document!;
	const rendered = renderPage(document, document.pages[0]!);
	const turned = rendered.svg.querySelector('[data-rotation3d]')!;
	expect(turned.hasAttribute('data-solid3d')).toBe(true);
	// Perspective cannot be a matrix: the outline is a projected polyline in the shape's own group.
	expect(turned.hasAttribute('transform')).toBe(false);
	const outline = turned.querySelector('[data-geometry]')!;
	expect(outline.getAttribute('d')).toMatch(/^M [-\d. ]+ L /);
	expect(turned.querySelector('[data-solid] > [data-contour]')).not.toBeNull();
	expect(turned.querySelector('[data-solid] > [data-extrusion]')).not.toBeNull();
	expect(turned.querySelector('[data-extrusion]')!.getAttribute('pointer-events')).toBe('none');
	const fills = [...turned.querySelectorAll('[data-extrusion] path')].map((p) =>
		p.getAttribute('fill'),
	);
	expect(fills.length).toBeGreaterThan(1);
	expect(fills.every((fill) => /^#[0-9a-f]{6}$/.test(fill!))).toBe(true);
	// The layers sit under the shape's own paths.
	expect(
		turned.querySelector('[data-solid]')!.compareDocumentPosition(outline) &
			Node.DOCUMENT_POSITION_FOLLOWING,
	).toBeTruthy();
	rendered.dispose();
	ui.dispose();
	ui.controller.destroy();
});

it('draws a depth only on a shape with a top bevel', async () => {
	const ui = await setup();
	ui.selection();
	ui.press('rotation-3d-isometric-right-up');
	await ui.done();
	ui.commands.run({ type: 'shape-format', patch: { extrusionDepth: 20 } });
	await ui.done();
	const document = ui.controller.state.document!;
	const rendered = renderPage(document, document.pages[0]!);
	expect(rendered.svg.querySelector('[data-extrusion]')).toBeNull();
	// Still the cheap parallel matrix: no perspective and nothing extruded.
	expect(rendered.svg.querySelector('[data-rotation3d]')!.getAttribute('transform')).toMatch(
		/^matrix\(/,
	);
	rendered.dispose();
	ui.dispose();
	ui.controller.destroy();
});
