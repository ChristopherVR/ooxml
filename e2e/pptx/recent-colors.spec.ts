/* oxlint-disable vitest/prefer-importing-vitest-globals -- Playwright spec, `test`/`expect` come from @playwright/test */
/**
 * "Recent colours": one MRU list, shown under every colour picker.
 *
 * PowerPoint remembers the last colours a user picked (`p:clrMru` in
 * `presentation.xml`) and offers them as a "Recent Colors" row in every colour
 * dropdown. The viewer models that list once per deck (shared
 * `render/recent-colors.ts`) and every binding is expected to:
 *
 *   - render the row (`data-testid="pptx-color-recent"`, the shared
 *     `pptx-ui-color-grid` with only Recent Colors) under the inspector's fill
 *     and stroke pickers, most-recent first, one button per colour carrying its
 *     hex in `data-color` and named "Recent #HEX";
 *   - hide the row while the list is empty;
 *   - push a colour into the list when a picker COMMITS (the `change` event,
 *     not the continuous `input` stream a native colour input emits while the
 *     user drags), so that a colour picked for a fill appears under the stroke
 *     picker too;
 *   - move a colour to the front when it is picked again, whether through a
 *     picker or by clicking its recent swatch.
 *
 * Each binding used to wire a different subset of pickers to its list, so the
 * feature was at parity nowhere. This spec drives the two pickers every
 * binding has (fill and stroke) and asserts the row contract through the same
 * `data-testid` in all five demos.
 *
 * The drag itself follows one shared rule too (`createLiveColorInput`): every
 * shade shows on the shape while the user drags, with no undo step, and the
 * settled pick is ONE undo step. Two bindings showed nothing until the picker
 * closed, and the other three recorded an undo step per shade.
 *
 * Run: bunx playwright test recent-colors
 */
import { expect, test } from '@playwright/test';
import type { Locator, Page } from '@playwright/test';

import { fixture, inspector, loadDeck } from './support/deck';

test.use({ viewport: { width: 1440, height: 900 } });

/** Two plain filled shapes ("SOURCE" / "TARGET"), no text-only elements. */
const DECK = fixture('format-painter.pptx');

const RED = '#ff0000';
const BLUE = '#0000ff';

/**
 * Every recent-colours row currently SHOWN inside the inspector. "Hidden while
 * empty" is satisfied two ways (four bindings unmount the row, vanilla keeps
 * it attached with `hidden` set), so only visibility is the contract.
 */
function recentRows(page: Page): Locator {
	return inspector(page).getByTestId('pptx-color-recent').filter({ visible: true });
}

/** The hexes one recent-colours row shows, most-recent first (the locator pierces its shadow root). */
async function rowColors(row: Locator): Promise<string[]> {
	return row
		.locator('button')
		.evaluateAll((buttons) =>
			buttons.map((button) => (button.getAttribute('data-color') ?? '').toLowerCase()),
		);
}

/**
 * The inspector's native colour inputs, in document order: for a shape the
 * fill picker comes first and the stroke picker after it in every binding.
 */
function colorInputs(page: Page): Locator {
	return inspector(page).locator('input[type="color"]:visible');
}

/**
 * Commit a colour through a native `<input type="color">` the way the picker
 * does when the user closes the swatch dialog: set the value and fire
 * `input` then `change`. The value is written through the prototype setter so
 * React's value tracker sees a real change rather than swallowing the event.
 */
async function pickColor(input: Locator, hex: string): Promise<void> {
	await input.evaluate((element, value) => {
		const setter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value')?.set;
		setter?.call(element, value);
		element.dispatchEvent(new Event('input', { bubbles: true }));
		element.dispatchEvent(new Event('change', { bubbles: true }));
	}, hex);
}

/**
 * Drag through the OS picker: one `input` event per shade passed, and no
 * `change`, which a real picker only fires once the pick settles. Each shade is
 * its own task, like the ticks of a real drag.
 */
async function dragColor(input: Locator, shades: readonly string[]): Promise<void> {
	for (const shade of shades) {
		await input.evaluate((element, value) => {
			const setter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value')?.set;
			setter?.call(element, value);
			element.dispatchEvent(new Event('input', { bubbles: true }));
		}, shade);
	}
}

/** The picker settles on `hex`: the one `change` a drag ends with. */
async function settleColor(input: Locator, hex: string): Promise<void> {
	await input.evaluate((element, value) => {
		const setter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value')?.set;
		setter?.call(element, value);
		element.dispatchEvent(new Event('change', { bubbles: true }));
	}, hex);
}

/** Two frames, so every binding has rendered whatever the last events changed. */
async function nextFrames(page: Page): Promise<void> {
	await page.evaluate(
		() =>
			new Promise<void>((resolve) => {
				requestAnimationFrame(() => requestAnimationFrame(() => resolve()));
			}),
	);
}

function sourceShape(page: Page): Locator {
	return page.locator('[data-pptx-element="true"]').filter({ hasText: 'SOURCE' });
}

async function selectSourceShape(page: Page): Promise<void> {
	const source = sourceShape(page);
	await source.waitFor();
	await source.click();
	await expect(colorInputs(page).first()).toBeVisible();
}

/**
 * The fill the shape paints: every binding writes the resolved fill onto the
 * element container's background (as `format-painter.spec.ts` reads it).
 */
async function shapeFill(shape: Locator): Promise<string> {
	return shape.evaluate((el) => getComputedStyle(el).backgroundColor);
}

/**
 * The ribbon's Undo control, by accessible name: it is a bare icon button in
 * every binding (as `inspector-undo.spec.ts` finds it).
 */
function undoButton(page: Page): Locator {
	return page.getByRole('button', { name: /^undo/iu }).first();
}

test.describe('recent colours', () => {
	test('the row is hidden until a picker commits, then lists the pick under fill AND stroke', async ({
		page,
	}) => {
		await loadDeck(page, DECK);
		await selectSourceShape(page);

		// The fixture authors no `p:clrMru`, so nothing has been picked yet.
		await expect(recentRows(page)).toHaveCount(0);

		await pickColor(colorInputs(page).first(), RED);

		// The fill picker's own row and the stroke picker's row read the same
		// list: a colour picked once is offered everywhere.
		await expect.poll(() => recentRows(page).count()).toBeGreaterThanOrEqual(2);
		const rows = recentRows(page);
		const count = await rows.count();
		for (let index = 0; index < count; index += 1) {
			const row = rows.nth(index);
			await expect(row).toBeVisible();
			expect(await rowColors(row)).toStrictEqual([RED]);
			await expect(row.locator('button').first()).toHaveAttribute(
				'aria-label',
				new RegExp(`^Recent ${RED}$`, 'iu'),
			);
		}
	});

	test('every recent-colours row stays inside the inspector', async ({ page }) => {
		await loadDeck(page, DECK);
		await selectSourceShape(page);
		await pickColor(colorInputs(page).first(), RED);
		await pickColor(colorInputs(page).nth(1), BLUE);
		await expect.poll(() => recentRows(page).count()).toBeGreaterThanOrEqual(2);

		// A row laid out beside its theme colours instead of under them ran past the panel's edge.
		const panel = await inspector(page).boundingBox();
		expect(panel).not.toBeNull();
		const rows = recentRows(page);
		const count = await rows.count();
		for (let index = 0; index < count; index += 1) {
			const box = await rows.nth(index).boundingBox();
			expect(box).not.toBeNull();
			expect(box!.x).toBeGreaterThanOrEqual(panel!.x - 1);
			expect(box!.x + box!.width).toBeLessThanOrEqual(panel!.x + panel!.width + 1);
		}
	});

	test('a second pick goes to the front, and re-picking a swatch moves it back to the front', async ({
		page,
	}) => {
		await loadDeck(page, DECK);
		await selectSourceShape(page);

		await pickColor(colorInputs(page).first(), RED);
		// Commit the second colour through the STROKE picker: the list is one
		// per deck, not one per picker.
		await pickColor(colorInputs(page).nth(1), BLUE);

		await expect.poll(async () => rowColors(recentRows(page).first())).toStrictEqual([BLUE, RED]);

		// Clicking the older swatch commits it again, so it becomes the most
		// recent pick and the row reorders rather than growing a duplicate.
		await recentRows(page).first().locator(`button[data-color="${RED}" i]`).click();
		await expect.poll(async () => rowColors(recentRows(page).first())).toStrictEqual([RED, BLUE]);
	});

	test('dragging through the fill picker adds nothing until the pick settles, then only that colour', async ({
		page,
	}) => {
		await loadDeck(page, DECK);
		await selectSourceShape(page);

		// The shades a drag passes on its way: none of them is a pick.
		const settled = '#cc0000';
		await dragColor(colorInputs(page).first(), ['#330000', '#660000', '#990000', settled]);
		await nextFrames(page);
		await expect(recentRows(page)).toHaveCount(0);

		// The picker settles where the drag ended: that one colour is the pick.
		await settleColor(colorInputs(page).first(), settled);
		await expect.poll(() => recentRows(page).count()).toBeGreaterThanOrEqual(1);
		const rows = recentRows(page);
		const count = await rows.count();
		for (let index = 0; index < count; index += 1) {
			expect(await rowColors(rows.nth(index))).toStrictEqual([settled]);
		}
	});

	test('dragging through the fill picker shows each shade on the shape before the pick settles', async ({
		page,
	}) => {
		await loadDeck(page, DECK);
		await selectSourceShape(page);
		const shape = sourceShape(page);

		// No `change` yet: the picker is still open, so this is all preview.
		for (const [shade, painted] of [
			['#330000', 'rgb(51, 0, 0)'],
			['#990000', 'rgb(153, 0, 0)'],
		] as const) {
			await dragColor(colorInputs(page).first(), [shade]);
			await expect.poll(() => shapeFill(shape)).toBe(painted);
		}
		// A preview is not an edit: Undo has nothing to take back yet.
		await expect(undoButton(page)).toBeDisabled();
	});

	test('a drag through the fill picker and its settling pick are one undo step', async ({
		page,
	}) => {
		await loadDeck(page, DECK);
		await selectSourceShape(page);
		const shape = sourceShape(page);
		const before = await shapeFill(shape);
		await expect(undoButton(page)).toBeDisabled();

		const settled = '#cc0000';
		await dragColor(colorInputs(page).first(), ['#330000', '#660000', '#990000', settled]);
		await settleColor(colorInputs(page).first(), settled);
		await expect.poll(() => shapeFill(shape)).toBe('rgb(204, 0, 0)');
		await expect(undoButton(page)).toBeEnabled();

		// One press takes the whole drag back, and leaves nothing more to undo. A
		// binding that recorded each shade lands on the shade before last instead.
		await undoButton(page).click();
		await expect.poll(() => shapeFill(shape)).toBe(before);
		await expect(undoButton(page)).toBeDisabled();
	});
});
