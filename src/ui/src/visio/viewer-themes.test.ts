import { afterEach, expect, it } from 'vitest';
import { visioBuiltInTheme, visioThemeVariantColors } from 'ooxml-core/visio';
import type { OfficeUiGallery } from '../ribbon/gallery';
import { setupFormattingViewer as setup } from './__fixtures__/formatting-viewer';

afterEach(() => document.body.replaceChildren());

const gallery = (root: ShadowRoot, id: string) =>
	root.querySelector<OfficeUiGallery>(`office-ui-gallery[data-menu="${id}"]`)!;
async function pick(root: ShadowRoot, id: string, item: string) {
	const element = gallery(root, id);
	element.open = true;
	await element.updateComplete;
	element.querySelector<HTMLButtonElement>(`[data-gallery-item="${item}"]`)!.click();
}

it('applies a built-in theme and its variants to the page, recolouring quick-styled shapes', async () => {
	const ui = await setup();
	const themes = gallery(ui.root, 'themes');
	const variants = gallery(ui.root, 'variants');
	expect(themes.hasAttribute('disabled')).toBe(false);
	expect(variants.hasAttribute('disabled')).toBe(true);
	expect(variants.title).toMatch(/Apply a theme/);
	expect(themes.state?.sections[0]?.items.map((item) => item.id).slice(0, 3)).toEqual([
		'theme-none',
		'theme-office',
		'theme-slate',
	]);
	expect(themes.state?.sections[0]?.items[0]?.applied).toBe(true);
	// A quick-styled shape, as dropped stencil masters are.
	ui.selection();
	await pick(ui.root, 'quick-styles', 'quick-style-2-4');
	await ui.done();
	await pick(ui.root, 'themes', 'theme-harbor');
	await ui.done();
	const page = ui.controller.state.document!.pages[0]!;
	expect(ui.edits.at(-1)).toEqual([{ type: 'set-page-theme', pageId: page.id, theme: 'harbor' }]);
	expect(page.theme).toMatchObject({ name: 'Harbor', builtIn: 'harbor', variant: 0 });
	expect(ui.shape().style.fill).toBe('#1f78b4');
	expect(themes.state?.sections[0]?.items.find((item) => item.applied)?.id).toBe('theme-harbor');
	expect(variants.hasAttribute('disabled')).toBe(false);
	expect(variants.state?.sections[0]?.items).toHaveLength(4);
	await pick(ui.root, 'variants', 'variant-1');
	await ui.done();
	expect(ui.edits.at(-1)).toEqual([{ type: 'set-page-theme', pageId: page.id, variant: 1 }]);
	expect(ui.controller.state.document!.pages[0]!.theme!.variant).toBe(1);
	expect(variants.state?.sections[0]?.items[1]?.applied).toBe(true);
	expect(visioThemeVariantColors(visioBuiltInTheme('harbor'), 1)).toHaveLength(7);
	await pick(ui.root, 'themes', 'theme-none');
	await ui.done();
	expect(ui.controller.state.document!.pages[0]!.theme).toBeUndefined();
	expect(variants.hasAttribute('disabled')).toBe(true);
	for (let step = 0; step < 4; step++) {
		ui.press('undo');
		await ui.done();
	}
	expect(ui.controller.exportVsdx().bytes).toEqual(ui.bytes);
	ui.dispose();
	ui.controller.destroy();
});

it('Colors in More Variants applies one theme colour scheme to the page and undoes', async () => {
	const ui = await setup();
	const more = ui.root.querySelector<HTMLElement>('[data-menu="variant-options"]')!;
	expect(more.hasAttribute('disabled')).toBe(false);
	ui.selection();
	await pick(ui.root, 'quick-styles', 'quick-style-2-4');
	await ui.done();
	await pick(ui.root, 'themes', 'theme-harbor');
	await ui.done();
	ui.root
		.querySelector<HTMLElement>('[data-menu="theme-colors"] [command="colors-ember"]')!
		.dispatchEvent(
			new CustomEvent('office-command', { bubbles: true, detail: { command: 'colors-ember' } }),
		);
	await ui.done();
	const page = ui.controller.state.document!.pages[0]!;
	expect(ui.edits.at(-1)).toEqual([{ type: 'set-page-theme', pageId: page.id, colors: 'ember' }]);
	expect(page.theme).toMatchObject({ builtIn: 'ember' });
	expect(ui.shape().style.fill).toBe('#d9480f');
	ui.press('undo');
	await ui.done();
	expect(ui.shape().style.fill).toBe('#1f78b4');
	ui.dispose();
	ui.controller.destroy();
});

const choose = async (ui: Awaited<ReturnType<typeof setup>>, menu: string, command: string) => {
	ui.root
		.querySelector<HTMLElement>(`[data-menu="${menu}"] [command="${command}"]`)!
		.dispatchEvent(new CustomEvent('office-command', { bubbles: true, detail: { command } }));
	await ui.done();
};
const checked = (root: ShadowRoot, menu: string) =>
	[...root.querySelectorAll<HTMLElement>(`[data-menu="${menu}"] office-ui-menu-item`)]
		.filter((item) => item.getAttribute('checked') === 'true')
		.map((item) => item.getAttribute('command'));

it('Fonts, Effects and Connectors apply one scheme each, mark it and undo', async () => {
	const ui = await setup();
	const page = () => ui.controller.state.document!.pages[0]!;
	const disabled = (menu: string) =>
		ui.root.querySelector(`[data-menu="${menu}"]`)!.hasAttribute('disabled');
	// Without a theme only Colors and Fonts apply; effect and connector styles belong to a theme.
	expect(disabled('theme-fonts')).toBe(false);
	expect(disabled('theme-effects')).toBe(true);
	expect(disabled('theme-connectors')).toBe(true);
	expect(checked(ui.root, 'theme-fonts')).toEqual([]);
	await choose(ui, 'theme-fonts', 'fonts-harbor');
	expect(ui.edits.at(-1)).toEqual([{ type: 'set-page-theme', pageId: page().id, fonts: 'harbor' }]);
	expect(page().themeSchemes).toEqual({ fonts: 'harbor' });
	expect(checked(ui.root, 'theme-fonts')).toEqual(['fonts-harbor']);
	expect(checked(ui.root, 'theme-colors')).toEqual([]);
	await pick(ui.root, 'themes', 'theme-slate');
	await ui.done();
	expect(disabled('theme-effects')).toBe(false);
	expect(disabled('theme-connectors')).toBe(false);
	await choose(ui, 'theme-effects', 'effects-ember');
	await choose(ui, 'theme-connectors', 'connectors-grove');
	expect(ui.edits.at(-1)).toEqual([
		{ type: 'set-page-theme', pageId: page().id, connectors: 'grove' },
	]);
	expect(page().themeSchemes).toMatchObject({
		colors: 'slate',
		fonts: 'slate',
		effects: 'ember',
		connectors: 'grove',
	});
	expect(checked(ui.root, 'theme-colors')).toEqual(['colors-slate']);
	expect(checked(ui.root, 'theme-effects')).toEqual(['effects-ember']);
	expect(checked(ui.root, 'theme-connectors')).toEqual(['connectors-grove']);
	ui.press('undo');
	await ui.done();
	expect(checked(ui.root, 'theme-connectors')).toEqual(['connectors-slate']);
	expect(checked(ui.root, 'theme-effects')).toEqual(['effects-ember']);
	ui.dispose();
	ui.controller.destroy();
});

it('Colors marks the current colour scheme', async () => {
	const ui = await setup();
	expect(checked(ui.root, 'theme-colors')).toEqual([]);
	await choose(ui, 'theme-colors', 'colors-ember');
	expect(checked(ui.root, 'theme-colors')).toEqual(['colors-ember']);
	await choose(ui, 'theme-colors', 'colors-harbor');
	expect(checked(ui.root, 'theme-colors')).toEqual(['colors-harbor']);
	ui.dispose();
	ui.controller.destroy();
});

it('Embellishment in More Variants writes the page level, with no theme needed', async () => {
	const ui = await setup();
	const page = () => ui.controller.state.document!.pages[0]!;
	await choose(ui, 'theme-embellishment', 'embellishment-high');
	expect(ui.edits.at(-1)).toEqual([
		{ type: 'set-page-theme', pageId: page().id, embellishment: 'high' },
	]);
	ui.dispose();
	ui.controller.destroy();
});
