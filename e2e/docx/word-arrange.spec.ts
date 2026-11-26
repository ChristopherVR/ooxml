import { expect, test } from '@playwright/test';
import { dialogByHeading } from './helpers';
import { choosePosition, openShapes, unfoldArrange } from './word-arrange-support';
import { savedDocument } from './saved-document';

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

test('page geometry commands explain why they are unavailable in Draft', async ({ page }) => {
	const editor = await openShapes(page);
	await editor.locator('.ProseMirror .dve-shape').first().click();
	await editor.getByRole('tab', { name: 'Layout', exact: true }).click();
	const align = editor.getByRole('combobox', { name: 'Align', exact: true });
	await unfoldArrange(editor, 'Layout');
	await expect(align).toBeDisabled();
	await expect(align).toHaveAccessibleDescription('Switch to Print Layout to arrange drawings.');
});

test('ungroup keeps the native children on the canvas in consecutive stacking order', async ({
	page,
}) => {
	const fixture = new URL(
		'../../src/core/docx/__fixtures__/shapes/arrange-group.docx',
		import.meta.url,
	);
	const editor = await openShapes(page, fixture);
	await editor.getByRole('button', { name: 'Print Layout', exact: true }).click();
	const floats = editor.locator('.dve-print-float');
	await expect(floats).toHaveCount(1);
	await editor.getByRole('tab', { name: 'Layout', exact: true }).click();
	await unfoldArrange(editor, 'Layout');
	await editor.getByRole('button', { name: 'Selection Pane', exact: true }).click();
	const pane = editor.getByRole('complementary', { name: 'Selection Pane' });
	await pane.getByRole('listitem').getByRole('button').click();
	await unfoldArrange(editor, 'Layout');
	await editor.getByRole('combobox', { name: 'Group', exact: true }).selectOption('ungroup');
	await expect(floats).toHaveCount(3);
	const { xml } = await savedDocument(page, 'ungrouped.docx');
	const heights = [...xml.matchAll(/relativeHeight="(\d+)"/g)].map((match) => Number(match[1]));
	expect(heights).toEqual([251659364, 251659365, 251659366]);
});

test('ordinary text edits leave unchanged Selection Pane rows mounted', async ({ page }) => {
	const editor = await openShapes(page);
	await editor.getByRole('tab', { name: 'Layout', exact: true }).click();
	await unfoldArrange(editor, 'Layout');
	await editor.getByRole('button', { name: 'Selection Pane', exact: true }).click();
	const row = editor
		.getByRole('complementary', { name: 'Selection Pane' })
		.getByRole('listitem')
		.first();
	await row.evaluate((element) => element.setAttribute('data-test-original-row', ''));
	const story = editor.locator('.ProseMirror').first();
	await story.press('Control+End');
	await story.press('End');
	await story.press('Space');
	await story.press('a');
	await expect(async () => {
		const button = row.getByRole('button');
		await button.click();
		await expect(button).toHaveAttribute('aria-pressed', 'true', { timeout: 500 });
	}).toPass();
	await expect
		.poll(() => row.evaluate((element) => element.hasAttribute('data-test-original-row')))
		.toBe(true);
});

test('Arrange controls fit the ribbon or its overflow in Shape Format and Layout', async ({
	page,
}) => {
	const editor = await openShapes(page);
	await editor.locator('.ProseMirror .dve-shape').first().click();
	for (const [tab, width] of [
		['Shape Format', 1800],
		['Layout', 2600],
		['Layout', 850],
		['Layout', 390],
	] as const) {
		await page.setViewportSize({ width, height: 950 });
		await editor.getByRole('tab', { name: tab, exact: true }).click();
		await unfoldArrange(editor, tab);
		const group = editor.locator('.ribbon-group[data-label="Arrange"]:visible').filter({
			has: page.getByRole('combobox', { name: 'Position', exact: true }),
		});
		await expect(group.locator(':is(button, select)[data-arrange]:visible')).toHaveCount(9);
		await expect
			.poll(() =>
				group.evaluate((element) => {
					const ribbon = element.closest('.ribbon-panel, .ribbon-overflow-panel')!;
					const bounds = ribbon.getBoundingClientRect();
					return [...element.querySelectorAll<HTMLElement>(':is(button, select)[data-arrange]')]
						.filter((control) => control.getClientRects().length)
						.filter((control) => {
							const { left, right, top, bottom } = control.getBoundingClientRect();
							return (
								left < bounds.left - 1 ||
								right > bounds.right + 1 ||
								top < bounds.top - 1 ||
								bottom > bounds.bottom + 1
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

test('Position presets place an ellipse at the top, center and bottom of the printed page', async ({
	page,
}) => {
	const editor = await openShapes(page);
	await editor.getByRole('button', { name: 'Print Layout', exact: true }).click();
	const ellipse = editor.locator('.dve-print-float').nth(2);
	await expect(ellipse).toBeVisible();
	await ellipse.click({ force: true });
	await editor.getByRole('tab', { name: 'Layout', exact: true }).click();
	const measure = () =>
		ellipse.evaluate((element) => {
			const bounds = element.getBoundingClientRect();
			const sheet = element.closest('.dve-print-page')!;
			const margins = sheet.querySelector('.dve-print-column')!.getBoundingClientRect();
			return {
				top: bounds.top - margins.top,
				right: margins.right - bounds.right,
				bottom: margins.bottom - bounds.bottom,
				centerX: bounds.left + bounds.width / 2 - margins.left - margins.width / 2,
				centerY: bounds.top + bounds.height / 2 - margins.top - margins.height / 2,
			};
		});
	for (const [preset, edges] of [
		['Top Right', ['top', 'right']],
		['Bottom Right', ['bottom', 'right']],
		['Middle Center', ['centerX', 'centerY']],
	] as const) {
		await choosePosition(editor, preset);
		for (const edge of edges)
			await expect.poll(async () => Math.abs((await measure())[edge])).toBeLessThan(1);
	}
});

test('pane and Shift-click selections align and distribute three drawings on the canvas', async ({
	page,
}) => {
	const editor = await openShapes(page);
	await editor.getByRole('button', { name: 'Print Layout', exact: true }).click();
	const floats = editor.locator('.dve-print-float');
	await expect(floats).toHaveCount(4);
	await editor.getByRole('tab', { name: 'Layout', exact: true }).click();
	await unfoldArrange(editor, 'Layout');
	await editor.getByRole('button', { name: 'Selection Pane', exact: true }).click();
	const pane = editor.getByRole('complementary', { name: 'Selection Pane' });
	const drawings = pane.getByRole('listitem').getByRole('button');
	await drawings.nth(1).click();
	await drawings.nth(2).click({ modifiers: ['Control'] });
	await drawings.nth(3).click({ modifiers: ['Control'] });
	await expect(pane.getByRole('button', { pressed: true })).toHaveCount(3);
	await expect(editor.locator('.dve-print-float-frame')).toHaveCount(3);
	const boxes = () =>
		floats.evaluateAll((elements) =>
			elements.slice(0, 3).map((element) => {
				const box = element.getBoundingClientRect();
				return { x: box.x, y: box.y, width: box.width, height: box.height };
			}),
		);
	const align = async (value: string) => {
		await unfoldArrange(editor, 'Layout');
		await editor.getByRole('combobox', { name: 'Align', exact: true }).selectOption(value);
	};
	for (const axis of ['x', 'y'] as const) {
		const size = axis === 'x' ? 'width' : 'height';
		const gapDelta = async () => {
			const [a, b, c] = (await boxes()).sort((left, right) => left[axis] - right[axis]);
			return Math.abs(b![axis] - a![axis] - a![size] - (c![axis] - b![axis] - b![size]));
		};
		expect(await gapDelta()).toBeGreaterThan(1);
		await align(axis);
		await expect.poll(gapDelta).toBeLessThan(1);
	}
	for (const [command, axis] of [
		['left', 'x'],
		['top', 'y'],
	] as const) {
		await align(command);
		await expect
			.poll(async () => {
				const values = (await boxes()).map((box) => box[axis]);
				return Math.max(...values) - Math.min(...values);
			})
			.toBeLessThan(1);
	}
	await expect(pane.getByRole('button', { pressed: true })).toHaveCount(3);
	await drawings.nth(3).click();
	await floats.nth(3).click({ force: true, modifiers: ['Shift'] });
	await expect(pane.getByRole('button', { pressed: true })).toHaveCount(2);
	await expect(editor.locator('.dve-print-float-frame')).toHaveCount(2);
});

test('Top Left moves a fresh canvas selection and a single undo restores its original bounds', async ({
	page,
}) => {
	const editor = await openShapes(page);
	await editor.getByRole('button', { name: 'Print Layout', exact: true }).click();
	const ellipse = editor.locator('.dve-print-float').nth(2);
	await expect(ellipse).toBeVisible();
	const offset = () =>
		ellipse.evaluate((element) => {
			const bounds = element.getBoundingClientRect();
			const sheet = element.closest('.dve-print-page')!.getBoundingClientRect();
			return { x: bounds.left - sheet.left, y: bounds.top - sheet.top };
		});
	const original = await offset();
	await ellipse.click({ force: true });
	await editor.getByRole('tab', { name: 'Layout', exact: true }).click();
	await choosePosition(editor, 'Top Left');
	await expect.poll(offset).not.toEqual(original);
	await choosePosition(editor, 'Top Left');
	await expect(editor.locator('.ProseMirror').first()).toBeFocused();
	await page.keyboard.press('Control+z');
	await expect.poll(offset).toEqual(original);
});

test('Selection Pane is docked and renames a drawing with F2 or a double-click', async ({
	page,
}) => {
	const editor = await openShapes(page);
	await editor.getByRole('tab', { name: 'Layout', exact: true }).click();
	await unfoldArrange(editor, 'Layout');
	await editor.getByRole('button', { name: 'Selection Pane', exact: true }).click();
	const pane = editor.getByRole('complementary', { name: 'Selection Pane' });
	const row = pane.getByRole('button', { name: 'Oval 3' });
	const dock = await pane.evaluate((element) => {
		const canvas = element.previousElementSibling!.getBoundingClientRect();
		const side = element.getBoundingClientRect();
		return [
			Math.abs(side.left - canvas.right),
			Math.abs(side.top - canvas.top),
			Math.abs(side.bottom - canvas.bottom),
		];
	});
	expect(Math.max(...dock)).toBeLessThan(1);
	await row.dblclick();
	const name = pane.getByRole('textbox', { name: 'Drawing name' });
	await name.fill('Diagram oval');
	await name.press('Enter');
	const renamed = pane.getByRole('button', { name: 'Diagram oval' });
	await expect(renamed).toBeFocused();
	await renamed.press('F2');
	await name.fill('Final oval');
	await name.press('Enter');
	const selected = pane.getByRole('button', { name: 'Final oval', pressed: true });
	const colors = () =>
		selected.evaluate((element) => {
			const style = getComputedStyle(element);
			return [style.color, style.backgroundColor];
		});
	await editor.evaluate((element) => Object.assign(element, { theme: 'light' }));
	const light = await colors();
	expect(light[0]).not.toBe(light[1]);
	await editor.evaluate((element) => Object.assign(element, { theme: 'dark' }));
	await expect.poll(colors).not.toEqual(light);
	const dark = await colors();
	expect(dark[0]).not.toBe(dark[1]);
	await page.emulateMedia({ forcedColors: 'active' });
	await expect(selected).toHaveCSS('forced-color-adjust', 'none');
	const { xml } = await savedDocument(page, 'renamed-drawing.docx');
	expect(xml).toContain('name="Final oval"');
	expect(xml).not.toContain('title="Final oval"');
});

test('Bring Forward and Send Backward change which drawing paints over an overlap', async ({
	page,
}) => {
	const editor = await openShapes(page);
	await editor.getByRole('button', { name: 'Print Layout', exact: true }).click();
	const floats = editor.locator('.dve-print-float');
	await expect(floats).toHaveCount(4);
	const key = (i: number) =>
		floats.nth(i).evaluate((element) => {
			const { blockId, floatIndex } = (element as HTMLElement).dataset;
			return `[data-block-id="${blockId}"][data-float-index="${floatIndex}"]`;
		});
	const rectangle = await key(0);
	const arrow = await key(1);
	await editor.getByRole('tab', { name: 'Layout', exact: true }).click();
	await unfoldArrange(editor, 'Layout');
	await editor.getByRole('button', { name: 'Selection Pane', exact: true }).click();
	const pane = editor.getByRole('complementary', { name: 'Selection Pane' });
	for (const name of ['Arrow: Right 2', 'Rectangle 1']) {
		await pane.getByRole('button', { name, exact: true }).click();
		await choosePosition(editor, 'Top Left');
	}
	const arrowFloat = floats.and(editor.locator(arrow));
	const top = () =>
		editor.evaluate(
			(host, { rectangle, arrow }) => {
				const root = host.shadowRoot!;
				const a = root.querySelector(`.dve-print-float${rectangle}`)?.getBoundingClientRect();
				const b = root.querySelector(`.dve-print-float${arrow}`)?.getBoundingClientRect();
				if (!a || !b) return undefined;
				const left = Math.max(a.left, b.left);
				const right = Math.min(a.right, b.right);
				const top = Math.max(a.top, b.top);
				const bottom = Math.min(a.bottom, b.bottom);
				if (left >= right || top >= bottom) return undefined;
				const hit = root
					.elementFromPoint((left + right) / 2, (top + bottom) / 2)
					?.closest('.dve-print-float');
				if (hit?.matches(arrow)) return 'arrow';
				if (hit?.matches(rectangle)) return 'rectangle';
				return undefined;
			},
			{ rectangle, arrow },
		);
	await expect(arrowFloat).toBeVisible();
	await expect.poll(top).toBe('arrow');
	await unfoldArrange(editor, 'Layout');
	await editor
		.getByRole('combobox', { name: 'Bring Forward', exact: true })
		.selectOption('forward');
	await expect.poll(top).toBe('rectangle');
	await unfoldArrange(editor, 'Layout');
	await editor
		.getByRole('combobox', { name: 'Send Backward', exact: true })
		.selectOption('backward');
	await expect.poll(top).toBe('arrow');
});

test('one floating shape disables Align and Group with accessible reasons', async ({ page }) => {
	const editor = await openShapes(page);
	await editor.getByRole('button', { name: 'Print Layout', exact: true }).click();
	await editor.locator('.dve-print-float').nth(2).click({ force: true });
	await editor.getByRole('tab', { name: 'Layout', exact: true }).click();
	await unfoldArrange(editor, 'Layout');
	const align = editor.getByRole('combobox', { name: 'Align', exact: true });
	const group = editor.getByRole('combobox', { name: 'Group', exact: true });
	await expect(align).toBeDisabled();
	await expect(align).toHaveAccessibleDescription('Select two or more floating drawings.');
	await expect(group).toBeDisabled();
	await expect(group).toHaveAccessibleDescription(
		'Select two or more floating shapes in one paragraph.',
	);
});

test('the print selection frame rotates around the shape center', async ({ page }) => {
	const editor = await openShapes(page);
	await editor.getByRole('button', { name: 'Print Layout', exact: true }).click();
	await editor.locator('.dve-print-float').nth(2).click({ force: true });
	const frame = editor.locator('.dve-print-float-frame');
	const starting = await frame.boundingBox();
	expect(starting).not.toBeNull();
	await editor.getByRole('tab', { name: 'Layout', exact: true }).click();
	await unfoldArrange(editor, 'Layout');
	await editor.getByRole('combobox', { name: 'Rotate', exact: true }).selectOption('right');
	await expect(frame).toHaveCSS('rotate', '90deg');
	await expect
		.poll(async () => {
			const rotated = await frame.boundingBox();
			if (!rotated || !starting) return Infinity;
			return Math.max(
				Math.abs(rotated.width - starting.height),
				Math.abs(rotated.height - starting.width),
			);
		})
		.toBeLessThan(1);
});
