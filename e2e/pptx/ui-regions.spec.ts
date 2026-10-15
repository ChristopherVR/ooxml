/* oxlint-disable vitest/prefer-importing-vitest-globals -- Playwright spec, `test`/`expect` come from @playwright/test */
/**
 * Host-supplied regions and capabilities, end to end.
 *
 * `customization.regions` hands a built-in region (the inspector, the Options
 * dialog body, a File tab page, ...) to the host as an empty container, and
 * `capabilities` switch whole abilities off. The decisions and the mount
 * lifecycle are shared (`render/customization/customization-regions.ts`); this
 * spec catches the failure the unit suites cannot see, a binding that never
 * renders the host container in the region's place.
 *
 * Renderers are functions, which a URL cannot carry, so each test installs
 * them through the live component handle every demo exposes as
 * `window.__pptxViewer` (`setRegion`, `addBackstagePage`, `setCapability`).
 *
 * Run: PPTX_E2E_PORT_OFFSET=600 bunx playwright test ui-regions --project=react
 */
import { expect, test } from '@playwright/test';
import type { Page } from '@playwright/test';

import {
	loadDeck,
	openRibbonTab,
	ribbon,
	SAMPLE_DECK,
	selectElement,
	slideElements,
} from './support/deck';
import { openOptionsDialog, optionsCategory } from './support/settings-dialog';

test.use({ viewport: { width: 1440, height: 900 } });

type RegionContext = {
	region: string;
	canEdit: boolean;
	selectedElementIds: readonly string[];
	setOption?: unknown;
};
type HandleWindow = {
	__pptxViewer?: Record<string, (...args: unknown[]) => void>;
};

async function handleReady(page: Page): Promise<void> {
	await page.waitForFunction(
		() => typeof (window as HandleWindow).__pptxViewer?.setRegion === 'function',
	);
}

/**
 * Replace `region` with a mount function that prints its context into the
 * container, so the test reads what the binding delivered.
 */
async function replaceWithProbe(page: Page, region: string): Promise<void> {
	await handleReady(page);
	await page.evaluate((id) => {
		const print = (el: HTMLElement, ctx: RegionContext) => {
			el.dataset.testRegion = ctx.region;
			el.dataset.testCanEdit = String(ctx.canEdit);
			el.dataset.testSelected = String(ctx.selectedElementIds.length);
			el.dataset.testHasSetOption = String(typeof ctx.setOption === 'function');
			el.textContent = `Host ${ctx.region}`;
		};
		(window as HandleWindow).__pptxViewer?.setRegion?.(
			id,
			(container: HTMLElement, context: RegionContext) => {
				const probe = document.createElement('div');
				probe.className = 'test-region-probe';
				container.append(probe);
				print(probe, context);
				return { update: (next: RegionContext) => print(probe, next) };
			},
		);
	}, region);
}

const probe = (page: Page, region: string) =>
	page.locator(`.test-region-probe[data-test-region="${region}"]`);

test.describe('UI regions and capabilities', () => {
	test('a host Settings view replaces the Options body inside the dialog frame', async ({
		page,
	}) => {
		await loadDeck(page, SAMPLE_DECK);
		await replaceWithProbe(page, 'dialog.options');
		const dialog = await openOptionsDialog(page, ['Options']);
		const host = dialog.locator('.test-region-probe');
		await expect(host).toHaveText('Host dialog.options');
		await expect(host).toHaveAttribute('data-test-has-set-option', 'true');
		await expect(optionsCategory(dialog, 'General')).toHaveCount(0);
		await page.keyboard.press('Escape');
		await expect(dialog).toBeHidden();
	});

	test('a host inspector fills the inspector slot and follows the selection', async ({ page }) => {
		await loadDeck(page, SAMPLE_DECK);
		await replaceWithProbe(page, 'inspector');
		const host = probe(page, 'inspector');
		await expect(host).toBeVisible();
		await expect(host).toHaveAttribute('data-test-selected', '0');
		await selectElement(page, slideElements(page).first());
		await expect(host).toHaveAttribute('data-test-selected', '1');
		await page.evaluate(() => (window as HandleWindow).__pptxViewer?.setRegion?.('inspector'));
		await expect(host).toHaveCount(0);
	});

	test('a host File tab page appears in the nav and mounts its body', async ({ page }) => {
		await loadDeck(page, SAMPLE_DECK);
		await handleReady(page);
		await page.evaluate(() =>
			(window as HandleWindow).__pptxViewer?.addBackstagePage?.({
				id: 'policies',
				label: 'Policies',
				render: (container: HTMLElement) => {
					container.textContent = 'Acme policies';
				},
			}),
		);
		await ribbon(page).getByRole('tab', { name: 'File', exact: true }).click();
		const backstage = page.locator('[role="dialog"][aria-label="File"]');
		await backstage.getByRole('button', { name: 'Policies', exact: true }).click();
		await expect(backstage.locator('[data-pptx-region="backstage.host.policies"]')).toHaveText(
			'Acme policies',
		);
	});

	test('capabilities.edit false makes the deck read-only, live', async ({ page }) => {
		await loadDeck(page, SAMPLE_DECK);
		await replaceWithProbe(page, 'statusBar');
		const host = probe(page, 'statusBar');
		await expect(host).toHaveAttribute('data-test-can-edit', 'true');
		await page.evaluate(() =>
			(window as HandleWindow).__pptxViewer?.setCapability?.('edit', false),
		);
		await expect(host).toHaveAttribute('data-test-can-edit', 'false');
		await page.evaluate(() =>
			(window as HandleWindow).__pptxViewer?.setCapability?.('edit', undefined),
		);
		await expect(host).toHaveAttribute('data-test-can-edit', 'true');
	});

	test('hiding a panel removes the ribbon controls that open it', async ({ page }) => {
		await loadDeck(page, SAMPLE_DECK);
		await openRibbonTab(page, 'View');
		const toggle = ribbon(page).locator('[data-ribbon-control="view.show.selectionPane"]');
		await expect(toggle).toBeVisible();
		await page.evaluate(() =>
			(window as HandleWindow).__pptxViewer?.setPanelVisible?.('selectionPane', false),
		);
		await expect(toggle).toBeHidden();
	});
});
