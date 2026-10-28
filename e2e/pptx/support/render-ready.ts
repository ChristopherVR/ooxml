import { expect, type Page } from '@playwright/test';

/** Let synchronous framework updates reach layout without an arbitrary wall-clock sleep. */
export async function nextPaint(page: Page): Promise<void> {
	await page.evaluate(
		() =>
			new Promise<void>((resolve) => {
				requestAnimationFrame(() => requestAnimationFrame(() => resolve()));
			}),
	);
}

/** The selected thumbnail is the shared contract, including an empty slide. */
export async function selectSlide(page: Page, slideNumber: number): Promise<void> {
	const thumbnail = page.locator(`[aria-label="Go to slide ${slideNumber}"]`).first();
	await thumbnail.click();
	await expect(thumbnail).toHaveAttribute('aria-current', /^(true|page)$/);
	await nextPaint(page);
}

export async function selectRibbonTab(page: Page, name: string): Promise<void> {
	const tab = page
		.getByRole('toolbar', { name: 'Presentation toolbar' })
		.getByRole('tab', { name, exact: true });
	await tab.click();
	await expect(tab).toHaveAttribute('aria-selected', 'true');
	await nextPaint(page);
}
