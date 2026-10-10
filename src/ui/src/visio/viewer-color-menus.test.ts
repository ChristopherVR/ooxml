import { afterEach, expect, it } from 'vitest';
import { parseVsdx } from 'ooxml-core/visio';
import { setupFormattingViewer as setup } from './__fixtures__/formatting-viewer';
import { colorAction, commonColor, pageThemeColors } from './ribbon-color-menu';

afterEach(() => document.body.replaceChildren());

type Ui = Awaited<ReturnType<typeof setup>>;
const moreColors = (ui: Ui) =>
	ui.root.querySelector<HTMLElement & { open: boolean }>('.more-colors-dialog')!;
const field = (ui: Ui, name: string) =>
	moreColors(ui).querySelector<HTMLInputElement>(`[data-color-field="${name}"]`)!;
const type = (input: HTMLInputElement, value: string) => {
	input.value = value;
	input.dispatchEvent(new Event('input', { bubbles: true }));
};
const ok = (ui: Ui) =>
	moreColors(ui)
		.querySelector('[command="more-colors-dialog-ok"]')!
		.shadowRoot!.querySelector('button')!
		.click();

it('maps picked colours to the formatting actions and reads the page theme', () => {
	expect(colorAction('fill', '#112233')).toEqual({
		type: 'shape-format',
		patch: { fillColor: '#112233' },
	});
	expect(colorAction('fill', 'none')).toEqual({
		type: 'shape-format',
		patch: { fillColor: 'none' },
	});
	expect(colorAction('line', 'none')).toEqual({ type: 'shape-format', patch: { linePattern: 0 } });
	expect(colorAction('line', '#112233')).toEqual({
		type: 'shape-format',
		patch: { lineColor: '#112233' },
	});
	expect(colorAction('font', '#112233')).toEqual({ type: 'font-color', value: '#112233' });
	// A drawing without a theme uses Visio's Office accents (the ones Quick Styles fall back to);
	// text and background are the palette's.
	const office = pageThemeColors(undefined);
	expect(office.slice(0, 4)).toEqual([undefined, undefined, undefined, undefined]);
	expect(office.slice(4).map((color) => color!.toLowerCase())).toEqual([
		'#5b9bd5',
		'#ed7d31',
		'#a5a5a5',
		'#ffc000',
		'#4472c4',
		'#70ad47',
	]);
	const themed = pageThemeColors({
		theme: { accents: ['#101010', '#202020', '#303030', '#404040', '#505050', '#606060'] },
	} as never);
	expect(themed.slice(4)).toEqual([
		'#101010',
		'#202020',
		'#303030',
		'#404040',
		'#505050',
		'#606060',
	]);
	expect(commonColor(['#AA0000', '#aa0000'])).toBe('#aa0000');
	expect(commonColor(['#aa0000', '#bb0000'])).toBeUndefined();
	expect(commonColor([])).toBeUndefined();
});

it('offers theme and standard colours in Fill, Line and Font Color and follows the selection', async () => {
	const ui = await setup();
	for (const target of ['fill', 'line', 'font'] as const) {
		const grid = ui.grid(target);
		expect(grid.closest('office-ui-menu-button')!.getAttribute('label')).toBe(
			{ fill: 'Fill', line: 'Line', font: 'Font Color' }[target],
		);
		expect(grid.shadowRoot!.querySelectorAll('[data-source="theme"]')).toHaveLength(60);
		expect(grid.shadowRoot!.querySelectorAll('[data-source="standard"]')).toHaveLength(10);
		expect(grid.shadowRoot!.querySelector('[data-command="more"]')!.textContent).toContain(
			'More Colors...',
		);
		expect(grid.value).toBeNull();
	}
	expect(ui.grid('fill').getAttribute('none-label')).toBe('No Fill');
	expect(ui.grid('line').getAttribute('none-label')).toBe('No Line');
	expect(ui.grid('font').hasAttribute('none-label')).toBe(false);
	ui.selection();
	// A theme tint, as Office names it.
	const tint = ui
		.grid('fill')
		.shadowRoot!.querySelector<HTMLButtonElement>('[aria-label="Accent 2, Lighter 60%"]')!;
	tint.click();
	await ui.done();
	expect(ui.edits.at(-1)).toEqual([
		{ type: 'format-shape', pageId: '1', shapeId: '1', fillColor: tint.dataset.color },
	]);
	expect(ui.shape().style.fill).toBe(tint.dataset.color);
	expect(ui.grid('fill').value).toBe(tint.dataset.color);
	ui.pickColor('fill', 'none');
	await ui.done();
	expect(ui.grid('fill').value).toBe('none');
	ui.dispose();
	ui.controller.destroy();
});

it('applies a custom colour from More Colors and lists it under Recent Colors', async () => {
	const ui = await setup();
	ui.selection();
	ui.pickColor('fill', '#ff0000');
	await ui.done();
	ui.pickColor('fill', 'more');
	expect(moreColors(ui).open).toBe(true);
	expect(moreColors(ui).getAttribute('heading')).toBe('Colors');
	// The dialog starts on the shape's colour.
	expect(field(ui, 'hex').value).toBe('#FF0000');
	expect(['red', 'green', 'blue'].map((name) => field(ui, name).value)).toEqual(['255', '0', '0']);
	type(field(ui, 'hex'), '#12ab');
	ok(ui);
	expect(moreColors(ui).querySelector('[role="alert"]')!.textContent).toMatch(/#RRGGBB/);
	expect(moreColors(ui).open).toBe(true);
	// Hex and the channels follow each other.
	type(field(ui, 'hex'), '1a2B3c');
	expect(['red', 'green', 'blue'].map((name) => field(ui, name).value)).toEqual(['26', '43', '60']);
	type(field(ui, 'blue'), '255');
	expect(field(ui, 'hex').value).toBe('#1A2BFF');
	type(field(ui, 'green'), '300');
	ok(ui);
	expect(moreColors(ui).open).toBe(true);
	type(field(ui, 'green'), '43');
	ok(ui);
	expect(moreColors(ui).open).toBe(false);
	await ui.done();
	expect(ui.shape().style.fill).toBe('#1a2bff');
	expect((await parseVsdx(ui.controller.exportVsdx().bytes)).pages[0]!.shapes[0]!.style.fill).toBe(
		'#1a2bff',
	);
	// Every picker now offers it again.
	for (const target of ['fill', 'line', 'font'] as const)
		expect(
			[...ui.grid(target).shadowRoot!.querySelectorAll<HTMLElement>('[data-source="recent"]')].map(
				(swatch) => swatch.dataset.color,
			),
		).toEqual(['#1a2bff']);
	// Cancel changes nothing.
	const count = ui.edits.length;
	ui.pickColor('line', 'more');
	type(field(ui, 'hex'), '#00ff00');
	moreColors(ui)
		.querySelector('[command="more-colors-dialog-cancel"]')!
		.shadowRoot!.querySelector('button')!
		.click();
	expect(moreColors(ui).open).toBe(false);
	expect(ui.edits).toHaveLength(count);
	ui.dispose();
	ui.controller.destroy();
});

it('sets the font colour from the grid and repeats it with the split button', async () => {
	const ui = await setup();
	ui.selection();
	ui.pickColor('font', '#7030a0');
	await ui.done();
	expect(ui.shape().text.color).toBe('#7030a0');
	expect(ui.grid('font').value).toBe('#7030a0');
	ui.press('undo');
	await ui.done();
	expect(ui.shape().text.color).not.toBe('#7030a0');
	ui.press('font-color');
	await ui.done();
	expect(ui.shape().text.color).toBe('#7030a0');
	ui.dispose();
	ui.controller.destroy();
});
