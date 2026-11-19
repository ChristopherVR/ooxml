import { expect, test, type Page } from '@playwright/test';
import { editor, newWorkbook, ribbon, sheetTabs } from './helpers';

type Probe = Window & { removedEditor?: WeakRef<HTMLElement> };

/** Forces collection until the probe's editor is gone. */
async function expectCollected(page: Page) {
	const session = await page.context().newCDPSession(page);
	await expect
		.poll(async () => {
			await session.send('HeapProfiler.collectGarbage');
			return page.evaluate(() => (window as Probe).removedEditor!.deref() === undefined);
		})
		.toBe(true);
}

const reconnect = (page: Page) =>
	editor(page).evaluate(async (element) => {
		const parent = element.parentElement!;
		element.remove();
		await new Promise((resolve) => setTimeout(resolve, 50));
		parent.append(element);
		await new Promise((resolve) => setTimeout(resolve, 100));
	});

test('a removed editor is collected', async ({ page }) => {
	await newWorkbook(page);
	await page.evaluate(async () => {
		const element = document.createElement('xlsx-editor');
		document.body.append(element);
		await new Promise((resolve) => setTimeout(resolve, 200));
		(window as Probe).removedEditor = new WeakRef(element);
		element.remove();
	});
	await expectCollected(page);
});

test('an editor removed with its context menu open is collected', async ({ page }) => {
	await newWorkbook(page);
	await editor(page).evaluate((element) => element.remove());
	await page.evaluate(async () => {
		const element = document.createElement('xlsx-editor') as HTMLElement & {
			newWorkbook(): void;
		};
		element.style.cssText = 'position:fixed;inset:0;z-index:9999';
		document.body.append(element);
		element.newWorkbook();
		await new Promise((resolve) => setTimeout(resolve, 200));
		(window as Probe).removedEditor = new WeakRef(element);
	});
	await sheetTabs(page).getByRole('tab').first().click({ button: 'right' });
	await expect(editor(page).locator('office-ui-context-menu')).toHaveCount(1);
	await editor(page).evaluate((element) => element.remove());
	await expectCollected(page);
});

test('a reconnected editor still folds its ribbon when the window narrows', async ({ page }) => {
	await page.setViewportSize({ width: 1800, height: 900 });
	await newWorkbook(page);
	await reconnect(page);
	const folded = ribbon(page).locator('.ribbon-panel:not([hidden]) .ribbon-group[data-collapsed]');
	await expect(folded).toHaveCount(0);
	await page.setViewportSize({ width: 700, height: 900 });
	await expect.poll(() => folded.count()).toBeGreaterThan(0);
});

test('a reconnected editor follows its language', async ({ page }) => {
	await newWorkbook(page);
	await reconnect(page);
	await editor(page).evaluate((element) => {
		(element as HTMLElement & { locale: string }).locale = 'fr';
	});
	await expect(ribbon(page).getByRole('tab', { name: 'Accueil' })).toBeVisible();
});
