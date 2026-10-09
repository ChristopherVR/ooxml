import { afterEach, expect, it } from 'vitest';
import { parseVsdx, visioFallbackQuickStyle } from 'ooxml-core/visio';
import { setupFormattingViewer as setup } from './__fixtures__/formatting-viewer';
import { renderPage } from './render-svg';

afterEach(() => document.body.replaceChildren());

it('applies glow, soft edge and reflection presets, checks them and draws them in SVG', async () => {
	const ui = await setup();
	expect(ui.button('glow-8-2').disabled).toBe(true);
	ui.selection();
	for (const id of ['glow-8-2', 'soft-edges-5', 'reflection-tight-0', 'glow-options'])
		expect(ui.button(id).disabled).toBe(false);
	expect(ui.button('glow-none').getAttribute('checked')).toBe('true');
	ui.press('glow-8-2');
	await ui.done();
	const accent2 = visioFallbackQuickStyle({ color: 3, matrix: 4 }).fill;
	expect(ui.edits.at(-1)).toEqual([
		{
			type: 'format-shape',
			pageId: '1',
			shapeId: '1',
			glow: { size: 8, color: accent2, transparency: 60 },
		},
	]);
	expect(ui.shape().style.glow).toMatchObject({ color: accent2 });
	expect(ui.button('glow-8-2').getAttribute('checked')).toBe('true');
	expect(ui.button('glow-none').getAttribute('checked')).toBe('false');
	ui.press('soft-edges-5');
	await ui.done();
	ui.press('reflection-half-4');
	await ui.done();
	expect(ui.button('soft-edges-5').getAttribute('checked')).toBe('true');
	expect(ui.button('reflection-half-4').getAttribute('checked')).toBe('true');
	const document = ui.controller.state.document!;
	const rendered = renderPage(document, document.pages[0]!);
	const glow = rendered.svg.querySelector('[data-glow]')!;
	expect(glow.getAttribute('opacity')).toBe('0.4');
	expect(glow.querySelector('path')!.getAttribute('fill')).toBe(accent2);
	const glowFilter = rendered.svg.querySelector(glow.getAttribute('filter')!.slice(4, -1))!;
	expect(glowFilter.querySelector('feMorphology')!.getAttribute('operator')).toBe('dilate');
	const reflection = rendered.svg.querySelector('[data-reflection]')!;
	expect(reflection.getAttribute('transform')).toMatch(/^matrix\(1 0 0 -1 0 -0\.05/);
	expect(reflection.getAttribute('mask')).toMatch(/^url\(#visio-reflection-\d+\)$/);
	const geometry = rendered.svg.querySelector('[data-geometry]')!;
	expect(geometry.getAttribute('filter')).toMatch(/^url\(#visio-soft-edges-\d+\)$/);
	// Effects sit beneath the shape's own geometry and are not hit-testable.
	expect(glow.nextElementSibling).toBe(geometry);
	expect(glow.getAttribute('pointer-events')).toBe('none');
	rendered.dispose();
	ui.press('glow-none');
	await ui.done();
	ui.press('soft-edges-0');
	await ui.done();
	ui.press('reflection-none');
	await ui.done();
	const saved = (await parseVsdx(ui.controller.exportVsdx().bytes)).pages[0]!.shapes[0]!.style;
	expect([saved.glow, saved.softEdges, saved.reflection]).toEqual([
		undefined,
		undefined,
		undefined,
	]);
	for (let step = 0; step < 6; step++) {
		ui.press('undo');
		await ui.done();
	}
	expect(ui.controller.exportVsdx().bytes).toEqual(ui.bytes);
	ui.dispose();
	ui.controller.destroy();
});

it('opens Format Shape from the Shape Styles launcher and applies changed effect values', async () => {
	const ui = await setup();
	const group = ui.root.querySelector('office-ui-ribbon-group[launcher="shape-styles-dialog"]')!;
	expect(group.hasAttribute('launcher-disabled')).toBe(true);
	ui.selection();
	expect(group.hasAttribute('launcher-disabled')).toBe(false);
	group.dispatchEvent(
		new CustomEvent('office-command', {
			detail: { command: 'shape-styles-dialog' },
			bubbles: true,
			composed: true,
		}),
	);
	const dialog = ui.root.querySelector<HTMLElement & { open: boolean }>('.format-shape-dialog')!;
	expect(dialog.open).toBe(true);
	const field = (name: string) =>
		dialog.querySelector<HTMLInputElement>(`[data-format-field="${name}"]`)!;
	expect(field('glowSize').value).toBe('0');
	expect(field('softEdges').value).toBe('0');
	field('glowSize').value = '11';
	field('glowColor').value = '#336699';
	field('glowTransparency').value = '25';
	field('softEdges').value = '2.5';
	dialog
		.querySelector('[command="format-shape-apply"]')!
		.dispatchEvent(new CustomEvent('office-command', { bubbles: true }));
	await ui.done();
	expect(ui.edits.at(-1)).toEqual([
		{
			type: 'format-shape',
			pageId: '1',
			shapeId: '1',
			glow: { size: 11, color: '#336699', transparency: 25 },
			softEdges: 2.5,
		},
	]);
	expect(ui.shape().style.softEdges).toBeCloseTo(2.5 / 72);
	expect(dialog.open).toBe(false);
	// Invalid values stay in the dialog with a message.
	group.dispatchEvent(
		new CustomEvent('office-command', {
			detail: { command: 'shape-styles-dialog' },
			bubbles: true,
			composed: true,
		}),
	);
	field('reflectionSize').value = '200';
	dialog
		.querySelector('[command="format-shape-apply"]')!
		.dispatchEvent(new CustomEvent('office-command', { bubbles: true }));
	expect(dialog.querySelector('[data-format-error]')!.textContent).toMatch(/reflection size/);
	expect(dialog.open).toBe(true);
	ui.dispose();
	ui.controller.destroy();
});
