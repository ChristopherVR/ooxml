/* oxlint-disable vitest/prefer-importing-vitest-globals -- Playwright spec, `test`/`expect` come from @playwright/test */
/**
 * The File page (backstage) belongs to the viewer, not the page around it: it fills the viewer
 * host's box, leaves a host sidebar next to the viewer visible and clickable, and Escape closes
 * it with focus back on the File tab, as Word, Excel and Visio do.
 */
import { expect, test } from '@playwright/test';
import type { Page } from '@playwright/test';

import { loadDeck, ribbonTab } from './support/deck';

test.use({ viewport: { width: 1440, height: 900 } });

const SIDEBAR_WIDTH = 160;

interface Box {
	left: number;
	top: number;
	right: number;
	bottom: number;
}

/**
 * Put the viewer next to a host sidebar, the way the Office suite mounts it, and return the
 * viewer root's box. The root is the outermost `[data-pptx-theme-root]`, which every binding emits.
 */
async function mountBesideSidebar(page: Page): Promise<Box> {
	return page.evaluate((width) => {
		let root = document.querySelector('[data-pptx-viewport]')?.closest('[data-pptx-theme-root]');
		while (root?.parentElement?.closest('[data-pptx-theme-root]')) {
			root = root.parentElement.closest('[data-pptx-theme-root]');
		}
		if (!(root instanceof HTMLElement) || !root.parentElement) {
			throw new Error('viewer root not found');
		}
		const shell = document.createElement('div');
		shell.style.cssText = 'display:flex;width:100vw;height:100vh';
		const sidebar = document.createElement('button');
		sidebar.id = 'e2e-host-sidebar';
		sidebar.textContent = 'Host sidebar';
		sidebar.style.cssText = `flex:none;width:${width}px`;
		const slot = document.createElement('div');
		slot.style.cssText = 'position:relative;display:flex;flex:1;min-width:0;height:100%';
		root.parentElement.insertBefore(shell, root);
		slot.append(root);
		shell.append(sidebar, slot);
		const box = root.getBoundingClientRect();
		return { left: box.left, top: box.top, right: box.right, bottom: box.bottom };
	}, SIDEBAR_WIDTH);
}

test('the File page stays inside the viewer and Escape returns focus to File', async ({ page }) => {
	await loadDeck(page);
	const host = await mountBesideSidebar(page);
	expect(host.left, 'the viewer sits beside the sidebar').toBeGreaterThanOrEqual(SIDEBAR_WIDTH - 1);

	const file = ribbonTab(page, 'File');
	await file.click();
	const backstage = page.getByRole('dialog', { name: 'File' });
	await expect(backstage).toBeVisible();

	const box = await backstage.boundingBox();
	expect(box, 'the File page is drawn').not.toBeNull();
	expect(box!.x).toBeGreaterThanOrEqual(host.left - 1);
	expect(box!.y).toBeGreaterThanOrEqual(host.top - 1);
	expect(box!.x + box!.width).toBeLessThanOrEqual(host.right + 1);
	expect(box!.y + box!.height).toBeLessThanOrEqual(host.bottom + 1);
	// The page covers the viewer it replaces, not a sliver of it.
	expect(box!.width).toBeGreaterThan((host.right - host.left) * 0.9);

	const sidebarOnTop = await page.evaluate(
		(width) => document.elementFromPoint(width / 2, 450)?.id === 'e2e-host-sidebar',
		SIDEBAR_WIDTH,
	);
	expect(sidebarOnTop, 'the host sidebar is not covered').toBeTruthy();

	await expect
		.poll(() => backstage.evaluate((panel) => panel.contains(document.activeElement)))
		.toBeTruthy();
	await page.keyboard.press('Escape');
	await expect(backstage).toBeHidden();
	await expect(file).toBeFocused();
});
