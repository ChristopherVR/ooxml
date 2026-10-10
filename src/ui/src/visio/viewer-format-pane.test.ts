import { afterEach, expect, it } from 'vitest';
import { parseVsdx } from 'ooxml-core/visio';
import type { OfficeUiColorGrid } from '../controls';
import { setupFormattingViewer as setup } from './__fixtures__/formatting-viewer';

afterEach(() => document.body.replaceChildren());

type Ui = Awaited<ReturnType<typeof setup>>;
const pane = (ui: Ui) => ui.root.querySelector<HTMLElement>('.format-pane')!;
const part = (ui: Ui, target: 'fill' | 'line') => {
	const set = pane(ui).querySelector<HTMLElement>(`[data-format-section="${target}"]`)!;
	return {
		none: set.querySelector<HTMLInputElement>('input[value="none"]')!,
		solid: set.querySelector<HTMLInputElement>('input[value="solid"]')!,
		color: set.querySelector<HTMLButtonElement>('.format-pane-color')!,
		grid: set.querySelector<OfficeUiColorGrid>('office-ui-color-grid')!,
		details: set.querySelector<HTMLElement>('.format-pane-details')!,
	};
};
const field = (ui: Ui, name: string) =>
	pane(ui).querySelector<HTMLInputElement | HTMLSelectElement>(`[data-pane-field="${name}"]`)!;
const change = (input: HTMLInputElement | HTMLSelectElement, value: string) => {
	input.value = value;
	input.dispatchEvent(new Event('change', { bubbles: true }));
};
const launch = (ui: Ui) =>
	ui.root.querySelector('office-ui-ribbon-group[launcher="shape-styles-dialog"]')!.dispatchEvent(
		new CustomEvent('office-command', {
			detail: { command: 'shape-styles-dialog' },
			bubbles: true,
			composed: true,
		}),
	);

it('opens from the Shape Styles launcher and explains an empty selection', async () => {
	const ui = await setup();
	const group = ui.root.querySelector('office-ui-ribbon-group[launcher="shape-styles-dialog"]')!;
	expect(group.hasAttribute('launcher-disabled')).toBe(false);
	launch(ui);
	// The task pane, not the old dialog.
	expect(ui.root.querySelector<HTMLElement & { open: boolean }>('.format-shape-dialog')!.open).toBe(
		false,
	);
	expect(pane(ui).dataset.paneView).toBe('format');
	expect(pane(ui).querySelector('.format-pane-hint')!.textContent).toBe(
		'Select a shape to format.',
	);
	expect(part(ui, 'fill').solid.disabled).toBe(true);
	expect(field(ui, 'lineWeight').disabled).toBe(true);
	ui.selection();
	expect(pane(ui).querySelector<HTMLElement>('.format-pane-hint')!.hidden).toBe(true);
	expect(part(ui, 'fill').solid.disabled).toBe(false);
	expect([...pane(ui).querySelectorAll('legend')].map((legend) => legend.textContent)).toEqual([
		'Fill',
		'Line',
	]);
	ui.dispose();
	ui.controller.destroy();
});

it('applies fill changes at once and follows the selection', async () => {
	const ui = await setup();
	ui.selection();
	const fill = part(ui, 'fill');
	expect(fill.solid.checked).toBe(true);
	// The colour button opens the shared grid with theme colours.
	expect(fill.grid.hidden).toBe(true);
	fill.color.click();
	expect(fill.grid.hidden).toBe(false);
	expect(fill.color.getAttribute('aria-expanded')).toBe('true');
	expect(fill.grid.hasAttribute('none-label')).toBe(false);
	fill.grid
		.shadowRoot!.querySelector<HTMLButtonElement>('[data-source="standard"][data-color="#ffc000"]')!
		.click();
	await ui.done();
	expect(ui.edits.at(-1)).toEqual([
		{ type: 'format-shape', pageId: '1', shapeId: '1', fillColor: '#ffc000' },
	]);
	expect(fill.grid.hidden).toBe(true);
	expect(fill.color.title).toBe('Fill color: #FFC000');
	change(field(ui, 'fillTransparency'), '40');
	await ui.done();
	expect(ui.edits.at(-1)).toEqual([
		{ type: 'format-shape', pageId: '1', shapeId: '1', fillTransparency: 40 },
	]);
	expect(field(ui, 'fillTransparency').value).toBe('40');
	// No fill hides the colour and transparency; Solid fill brings the colour back.
	fill.none.checked = true;
	fill.none.dispatchEvent(new Event('change', { bubbles: true }));
	await ui.done();
	expect(ui.shape().style.fill).toBe('none');
	expect(fill.none.checked).toBe(true);
	expect(fill.details.hidden).toBe(true);
	fill.solid.checked = true;
	fill.solid.dispatchEvent(new Event('change', { bubbles: true }));
	await ui.done();
	expect(ui.shape().style.fill).toBe('#ffc000');
	expect(fill.details.hidden).toBe(false);
	// The ribbon's Fill picker shows the same colour.
	expect(ui.grid('fill').value).toBe('#ffc000');
	// One step per change.
	ui.press('undo');
	await ui.done();
	expect(ui.shape().style.fill).toBe('none');
	expect(fill.none.checked).toBe(true);
	ui.dispose();
	ui.controller.destroy();
});

it('sets line colour, width, dash type and transparency, and refuses bad numbers', async () => {
	const ui = await setup();
	ui.selection();
	const line = part(ui, 'line');
	line.color.click();
	line.grid.shadowRoot!.querySelector<HTMLButtonElement>('[data-color="#0070c0"]')!.click();
	await ui.done();
	change(field(ui, 'lineWeight'), '2.25');
	await ui.done();
	change(field(ui, 'linePattern'), '4');
	await ui.done();
	change(field(ui, 'lineTransparency'), '15');
	await ui.done();
	expect(ui.edits.slice(-4).map((edits) => edits[0])).toEqual([
		{ type: 'format-shape', pageId: '1', shapeId: '1', lineColor: '#0070c0' },
		{ type: 'format-shape', pageId: '1', shapeId: '1', lineWeight: 2.25 },
		{ type: 'format-shape', pageId: '1', shapeId: '1', linePattern: 4 },
		{ type: 'format-shape', pageId: '1', shapeId: '1', lineTransparency: 15 },
	]);
	expect(field(ui, 'lineWeight').value).toBe('2.25');
	expect(field(ui, 'linePattern').value).toBe('4');
	const saved = (await parseVsdx(ui.controller.exportVsdx().bytes)).pages[0]!.shapes[0]!;
	expect(saved.style.lineColor).toBe('#0070c0');
	expect(saved.style.lineWidth * 72).toBeCloseTo(2.25);
	const count = ui.edits.length;
	change(field(ui, 'lineWeight'), '999');
	expect(ui.edits).toHaveLength(count);
	expect(ui.feedback.at(-1)).toBe('Enter a line width from 0 to 150.');
	expect(field(ui, 'lineWeight').value).toBe('2.25');
	// No line and back.
	line.none.checked = true;
	line.none.dispatchEvent(new Event('change', { bubbles: true }));
	await ui.done();
	expect(ui.shape().style.linePattern).toBe(0);
	expect(line.details.hidden).toBe(true);
	line.solid.checked = true;
	line.solid.dispatchEvent(new Event('change', { bubbles: true }));
	await ui.done();
	expect(ui.shape().style.linePattern).toBe(1);
	expect(line.solid.checked).toBe(true);
	ui.dispose();
	ui.controller.destroy();
});

it('reaches the effects and pattern dialogs and takes a custom colour', async () => {
	const ui = await setup();
	ui.selection();
	pane(ui).querySelector<HTMLButtonElement>('[data-format-link="effects"]')!.click();
	const effects = ui.root.querySelector<HTMLElement & { open: boolean; close(): void }>(
		'.format-shape-dialog',
	)!;
	expect(effects.open).toBe(true);
	effects.close();
	pane(ui).querySelector<HTMLButtonElement>('[data-format-link="patterns"]')!.click();
	const patterns = ui.root.querySelector<HTMLElement & { open: boolean; close(): void }>(
		'.paint-properties-dialog:not(.format-shape-dialog)',
	)!;
	expect(patterns.open).toBe(true);
	patterns.close();
	const fill = part(ui, 'fill');
	fill.color.click();
	fill.grid.shadowRoot!.querySelector<HTMLButtonElement>('[data-command="more"]')!.click();
	const colors = ui.root.querySelector<HTMLElement & { open: boolean }>('.more-colors-dialog')!;
	expect(colors.open).toBe(true);
	const hex = colors.querySelector<HTMLInputElement>('[data-color-field="hex"]')!;
	hex.value = '#abcdef';
	hex.dispatchEvent(new Event('input', { bubbles: true }));
	colors
		.querySelector('[command="more-colors-dialog-ok"]')!
		.shadowRoot!.querySelector('button')!
		.click();
	await ui.done();
	expect(ui.shape().style.fill).toBe('#abcdef');
	expect(fill.grid.shadowRoot!.querySelector('[data-source="recent"]')).not.toBeNull();
	ui.dispose();
	ui.controller.destroy();
});
