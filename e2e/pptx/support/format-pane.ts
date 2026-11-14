/**
 * The inspector's Format pane sections (`pptx-ui-format-section`), found by
 * their framework-neutral `data-pptx-format-section` id.
 */
import { expect } from '@playwright/test';
import type { Locator, Page } from '@playwright/test';

import { inspector } from './deck';

/** One Format pane section of the visible inspector (`fill`, `sizePosition`, ...). */
export function formatSection(page: Page, id: string): Locator {
	return inspector(page).locator(`[data-pptx-format-section="${id}"]`).first();
}

/**
 * Expand a Format pane section when it is closed, so a spec can reach the
 * controls inside it (sections other than the selection's first ones start
 * closed).
 */
export async function openFormatSection(page: Page, id: string): Promise<Locator> {
	const section = formatSection(page, id);
	await section.waitFor();
	if ((await section.getAttribute('open')) === null) {
		await section.locator('summary').click();
	}
	await expect(section).toHaveAttribute('open', '');
	return section;
}
