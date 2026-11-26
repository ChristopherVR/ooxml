import { expect, test } from '@playwright/test';
import { dialogByHeading } from './helpers';
import { openShapes, unfoldArrange } from './word-arrange-support';
import { savedDocument } from './saved-document';

for (const framework of ['vanilla', 'react', 'vue', 'angular', 'svelte', 'solid']) {
	test(`${framework}: Wrap Text moves a floating shape behind the text and saves the anchor`, async ({
		page,
	}) => {
		const editor = await openShapes(page, { framework });
		await editor.getByRole('button', { name: 'Print Layout', exact: true }).click();
		await unfoldArrange(editor, 'Layout');
		await editor.getByRole('button', { name: 'Selection Pane', exact: true }).click();
		const pane = editor.getByRole('complementary', { name: 'Selection Pane' });
		await pane.getByRole('button', { name: 'Oval 3', exact: true }).click();
		await pane.getByRole('button', { name: 'Close', exact: true }).click();
		const wrap = editor.getByRole('combobox', { name: 'Wrap text', exact: true });
		const behind = editor.locator('.dve-print-float-behind');
		const before = await behind.count();
		await wrap.selectOption('behind');
		await expect(behind).toHaveCount(before + 1);
		const { xml } = await savedDocument(page, `wrapped-drawing-${framework}.docx`);
		const ovalAnchor = xml.match(
			/<wp:anchor\b[^>]*>(?:(?!<\/wp:anchor>)[\s\S])*?name="Oval 3"(?:(?!<\/wp:anchor>)[\s\S])*?<\/wp:anchor>/,
		)?.[0];
		expect(ovalAnchor).toContain('behindDoc="1"');
		expect(ovalAnchor).toContain('<wp:wrapNone');
	});
}

test('Word arranges a floating shape from its Layout ribbon and saves the edit', async ({
	page,
}) => {
	const editor = await openShapes(page);
	await editor.getByRole('button', { name: 'Print Layout', exact: true }).click();
	const shape = editor.locator('.dve-print-float').first();
	await expect(shape).toBeVisible();
	await editor.getByRole('tab', { name: 'Layout', exact: true }).click();
	const altText = editor.getByRole('button', { name: 'Alt Text', exact: true });
	await unfoldArrange(editor, 'Layout');
	await editor.getByRole('button', { name: 'Selection Pane', exact: true }).click();
	const pane = editor.getByRole('complementary', { name: 'Selection Pane' });
	await expect(pane.getByRole('listitem')).toHaveCount(4);
	await pane.getByRole('listitem').first().getByRole('button').click();
	await expect(pane.getByRole('button', { pressed: true })).toHaveCount(1);
	await pane.getByRole('button', { name: 'Close', exact: true }).click();
	await editor.getByRole('tab', { name: 'Layout', exact: true }).click();
	await unfoldArrange(editor, 'Layout');
	const position = editor.getByRole('combobox', { name: 'Position', exact: true });
	await position.press('Enter');
	await editor.getByRole('menuitemradio', { name: /Bottom Right/ }).click();
	await unfoldArrange(editor, 'Layout');
	await editor.getByRole('combobox', { name: 'Rotate', exact: true }).selectOption('right');
	await unfoldArrange(editor, 'Layout');
	await altText.click();
	const alt = dialogByHeading(editor, 'Alt Text');
	await alt.getByRole('textbox', { name: 'Description', exact: true }).fill('Arrangement sample');
	await alt.getByRole('button', { name: 'OK', exact: true }).click();
	const { xml } = await savedDocument(page, 'word-arrange.docx');
	expect(xml).toContain('<wp:positionH relativeFrom="margin"><wp:align>right</wp:align>');
	expect(xml).toContain('<wp:positionV relativeFrom="margin"><wp:align>bottom</wp:align>');
	expect(xml).toContain('descr="Arrangement sample"');
	expect(xml).toMatch(/<a:xfrm[^>]*rot="5400000"/);
});

test('Position explains why it is unavailable in Draft and Web Layout', async ({ page }) => {
	const editor = await openShapes(page);
	await editor.locator('.ProseMirror .dve-shape').first().click();
	const align = editor.getByRole('combobox', { name: 'Align', exact: true });
	const position = editor.getByRole('combobox', { name: 'Position', exact: true });
	const note = editor
		.getByRole('tabpanel', { name: 'Layout', exact: true })
		.getByText('Print Layout only', { exact: true });
	const unavailable = async () => {
		await expect(align).toBeDisabled();
		await expect(position).toBeDisabled();
		await expect(position).toHaveAccessibleDescription(
			'Switch to Print Layout to arrange drawings.',
		);
		await expect(note).toBeVisible();
	};
	await unfoldArrange(editor, 'Layout');
	await unavailable();
	await editor.getByRole('button', { name: 'Print Layout', exact: true }).click();
	await expect(position).toBeEnabled();
	await expect(note).toBeHidden();
	await editor.getByRole('button', { name: 'Web Layout', exact: true }).click();
	await unavailable();
});

test('Arrange controls fit the ribbon or its overflow in Shape Format and Layout', async ({
	page,
}) => {
	const editor = await openShapes(page);
	await editor.locator('.ProseMirror .dve-shape').first().click();
	for (const [tab, width] of [
		['Shape Format', 2000],
		['Shape Format', 1600],
		['Layout', 2000],
		['Layout', 1600],
		['Layout', 850],
		['Layout', 390],
	] as const) {
		await page.setViewportSize({ width, height: 950 });
		await editor.getByRole('tab', { name: tab, exact: true }).click();
		if (width >= 1600) {
			const panel = editor.getByRole('tabpanel', { name: tab, exact: true });
			await expect(panel.locator('.ribbon-group[data-label="Arrange"]')).not.toHaveAttribute(
				'data-collapsed',
				'',
			);
		} else await unfoldArrange(editor, tab);
		const group = editor.locator('.ribbon-group[data-label="Arrange"]:visible').filter({
			has: page.getByRole('combobox', { name: 'Position', exact: true }),
		});
		await expect(group.locator(':is(button, select)[data-arrange]:visible')).toHaveCount(9);
		await expect
			.poll(() =>
				group.evaluate((element) => {
					const ribbon = element.closest('.ribbon-panel, .ribbon-overflow-panel')!;
					const bounds = ribbon.getBoundingClientRect();
					const groupBottom = element.getBoundingClientRect().bottom;
					return [...element.querySelectorAll<HTMLElement>(':is(button, select)[data-arrange]')]
						.filter((control) => control.getClientRects().length)
						.filter((control) => {
							const { left, right, top, bottom } = control.getBoundingClientRect();
							return (
								left < bounds.left - 1 ||
								right > bounds.right + 1 ||
								top < bounds.top - 1 ||
								bottom > Math.min(bounds.bottom, groupBottom - 16) + 1
							);
						})
						.map((control) => ({
							label: control.getAttribute('aria-label'),
							control: control.getBoundingClientRect().toJSON(),
							wrapper: control.parentElement!.getBoundingClientRect().toJSON(),
							group: element.getBoundingClientRect().toJSON(),
							ribbon: bounds.toJSON(),
						}));
				}),
			)
			.toEqual([]);
	}
});
