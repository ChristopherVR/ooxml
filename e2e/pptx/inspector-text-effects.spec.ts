/* oxlint-disable vitest/prefer-importing-vitest-globals -- Playwright spec, `test`/`expect` come from @playwright/test */
/**
 * The format inspector's Text Effects: text shadow, glow and reflection of the TEXT
 * (`a:rPr/a:effectLst`), as opposed to the shape's own Effects. Every binding offers them as
 * checkboxes in `<label>`s captioned "Text Shadow", "Text Glow" and "Text Reflection" (React and
 * Svelte with their own controls, Vue, Angular and vanilla through the shared
 * `TEXT_EFFECT_SECTIONS`).
 *
 * Two bindings had none on screen: Vue imported its panels and never rendered them, and Angular
 * had 3-D text but no shadow, glow or reflection. And Svelte and vanilla (and Vue's unrendered
 * panel) sent a change to the body style alone, while a run is painted from its own style: the
 * box ticked and the slide showed nothing. So this checks the effect ON THE SLIDE (the run's CSS
 * `filter` / `text-shadow`), that it is one undo step, and that a save writes it into the run's
 * `a:rPr` as `a:glow` / `a:outerShdw`.
 *
 * Some bindings keep the section in a collapsed `<details>`, so the spec opens every inspector
 * section first, as a user would, and then finds the controls by their captions.
 *
 * Run: bunx playwright test inspector-text-effects
 */
import { expect, test } from '@playwright/test';
import type { Locator, Page } from '@playwright/test';
import JSZip from 'jszip';

import { savePptxViaBackstage } from './save-pptx';
import { elementWithText, inspector, loadDeck, selectElement } from './support/deck';
import { downloadBytes } from './support/exports';

test.use({ viewport: { width: 1440, height: 900 } });

/** A plain text box on slide 1 of the sample deck: one run with its own `a:rPr` (size, colour). */
const TEXT = 'Product Overview';

/** Expand every collapsed inspector section, as a user opening them would. */
async function openInspectorSections(page: Page): Promise<void> {
	await inspector(page).evaluate((root) => {
		for (const details of root.querySelectorAll('details')) {
			details.open = true;
		}
	});
}

/**
 * One of the Text Effects checkboxes: the checkbox of the `<label>` captioned with the name every
 * binding gives it. The checkbox is `<pptx-ui-checkbox>`, a form-associated custom element that
 * takes its accessible name from that label (Chromium's accessibility tree reads "Text Glow"), but
 * Playwright's own role and label matching only follow a `<label>` to native controls, so
 * `getByRole('checkbox', { name })` finds it only where a binding also sets `aria-label`.
 */
function effectToggle(page: Page, name: 'Text Shadow' | 'Text Glow' | 'Text Reflection'): Locator {
	return inspector(page)
		.locator('label')
		.filter({ hasText: new RegExp(`^\\s*${name}\\s*$`, 'u') })
		.getByRole('checkbox');
}

/** The ribbon's Undo control (an icon button: its name is the one thing all five agree on). */
function undoButton(page: Page): Locator {
	return page.getByRole('button', { name: /^undo/iu }).first();
}

/**
 * What the slide paints for the text box's runs: whether any element holding its text carries a
 * glow (`filter: drop-shadow(...)`) or a text shadow, read from the computed style.
 */
function paintedEffects(page: Page): Promise<{ glow: boolean; shadow: boolean }> {
	return elementWithText(page, TEXT).evaluate((root) => {
		const carriers = [root, ...root.querySelectorAll<HTMLElement>('*')].filter((node) =>
			[...node.childNodes].some(
				(child) => child.nodeType === Node.TEXT_NODE && (child.textContent ?? '').trim() !== '',
			),
		);
		// A run's effect may sit on the text's own span or on a wrapper around it.
		const styled = (matches: (style: CSSStyleDeclaration) => boolean): boolean =>
			carriers.some((node) => {
				for (
					let el: Element | null = node;
					el && el !== root.parentElement;
					el = el.parentElement
				) {
					if (matches(getComputedStyle(el))) {
						return true;
					}
				}
				return false;
			});
		return {
			glow: styled((style) => style.filter.includes('drop-shadow')),
			shadow: styled((style) => style.textShadow !== '' && style.textShadow !== 'none'),
		};
	});
}

async function selectTextBox(page: Page): Promise<void> {
	await loadDeck(page);
	await selectElement(page, elementWithText(page, TEXT));
	await expect(inspector(page)).toBeVisible();
	await openInspectorSections(page);
	await expect(effectToggle(page, 'Text Glow')).toBeVisible();
}

/** The saved `a:r` that carries the text box's text. */
async function savedRun(page: Page): Promise<string> {
	const bytes = await downloadBytes(await savePptxViaBackstage(page));
	const zip = await JSZip.loadAsync(bytes);
	const xml = await zip.file('ppt/slides/slide1.xml')!.async('string');
	const run = xml.split(/<a:r>/u).find((chunk) => chunk.includes(`>${TEXT}</a:t>`));
	expect(run, `no saved run holds "${TEXT}"`).toBeDefined();
	return run!.slice(0, run!.indexOf('</a:r>'));
}

test.describe('inspector text effects', () => {
	test('text glow and shadow paint on the slide and save into the run', async ({ page }) => {
		await selectTextBox(page);
		expect(await paintedEffects(page)).toStrictEqual({ glow: false, shadow: false });

		const glow = effectToggle(page, 'Text Glow');
		await expect(glow).not.toBeChecked();
		await glow.click();
		await expect(glow).toBeChecked();
		await expect.poll(() => paintedEffects(page)).toStrictEqual({ glow: true, shadow: false });

		const shadow = effectToggle(page, 'Text Shadow');
		await shadow.click();
		await expect(shadow).toBeChecked();
		await expect.poll(() => paintedEffects(page)).toStrictEqual({ glow: true, shadow: true });

		const run = await savedRun(page);
		expect(run).toMatch(/<a:effectLst>/u);
		expect(run).toMatch(/<a:glow\b[^>]*\brad="\d+"/u);
		expect(run).toMatch(/<a:outerShdw\b/u);
	});

	test('turning text glow on is one undo step', async ({ page }) => {
		await selectTextBox(page);
		await expect(undoButton(page)).toBeDisabled();

		await effectToggle(page, 'Text Glow').click();
		await expect.poll(() => paintedEffects(page)).toStrictEqual({ glow: true, shadow: false });
		await expect(undoButton(page)).toBeEnabled();

		await undoButton(page).click();
		await expect.poll(() => paintedEffects(page)).toStrictEqual({ glow: false, shadow: false });
	});

	test('text reflection has a toggle that turns on', async ({ page }) => {
		await selectTextBox(page);
		const reflection = effectToggle(page, 'Text Reflection');
		await expect(reflection).not.toBeChecked();
		await reflection.click();
		await expect(reflection).toBeChecked();
	});
});
