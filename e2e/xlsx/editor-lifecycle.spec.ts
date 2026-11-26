import { expect, test, type Locator, type Page } from '@playwright/test';
import { openFixture } from './fixture-files';
import { editor, newWorkbook, ribbon, sheetTabs } from './helpers';

type Probe = Window & { removedEditor?: WeakRef<HTMLElement> };
type Workbook = {
	sheets: { rows: Map<number, Map<number, { value: unknown }>>; autoFilter?: unknown }[];
};
type Editor = HTMLElement & {
	newWorkbook(): void;
	workbook: Workbook;
	collaboration: unknown;
	collaborationState: { active: boolean; people: unknown[] };
};

/** An editor of the probe's own (the demo page holds on to its editor), in a fixed host `#id`. */
async function addEditor(page: Page, id: string, box: string) {
	await page.evaluate(
		([hostId, css]) => {
			const host = document.createElement('div');
			host.id = hostId!;
			host.style.cssText = `position:fixed;z-index:9999;${css}`;
			const element = document.createElement('xlsx-editor') as Editor;
			element.style.cssText = 'display:block;width:100%;height:100%';
			host.append(element);
			document.body.append(host);
			element.newWorkbook();
		},
		[id, box],
	);
	const probe = page.locator(`#${id} xlsx-editor`);
	await expect(probe.locator('[part="grid"]')).toBeVisible();
	return probe;
}

const removeProbe = (probe: Locator) =>
	probe.evaluate((element) => {
		(window as Probe).removedEditor = new WeakRef(element as HTMLElement);
		element.remove();
	});

/** Capture-phase `pointerdown` listeners on the document, which keep an open popup's editor alive. */
async function documentPointerListeners(page: Page): Promise<number> {
	const session = await page.context().newCDPSession(page);
	const { result } = await session.send('Runtime.evaluate', { expression: 'document' });
	const { listeners } = await session.send('DOMDebugger.getEventListeners', {
		objectId: result.objectId!,
	});
	await session.detach();
	return listeners.filter((listener) => listener.type === 'pointerdown' && listener.useCapture)
		.length;
}

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
		(window as Probe).removedEditor = new WeakRef(element as HTMLElement);
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
		(window as Probe).removedEditor = new WeakRef(element as HTMLElement);
	});
	await sheetTabs(page).getByRole('tab').first().click({ button: 'right' });
	await expect(editor(page).locator('office-ui-context-menu')).toHaveCount(1);
	await editor(page).evaluate((element) => element.remove());
	await expectCollected(page);
});

test('a reconnected editor still folds its ribbon when the window narrows', async ({ page }) => {
	await page.setViewportSize({ width: 2300, height: 900 });
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

test('an editor removed with its filter menu open is collected', async ({ page }) => {
	await newWorkbook(page);
	const probe = await addEditor(page, 'probe', 'inset:0');
	await probe.evaluate((element) => {
		const excel = element as Editor;
		const workbook = excel.workbook;
		const sheet = workbook.sheets[0]!;
		[['Kind'], ['apple'], ['pear']].forEach((values, row) =>
			sheet.rows.set(row, new Map(values.map((value, col) => [col, { value }]))),
		);
		sheet.autoFilter = { range: { start: { row: 0, col: 0 }, end: { row: 2, col: 0 } } };
		excel.workbook = workbook;
	});
	const before = await documentPointerListeners(page);
	await probe.locator('[part="grid"] .xg-fb').first().click();
	await expect(probe.locator('.xg-fmenu')).toHaveCount(1);
	expect(await documentPointerListeners(page)).toBeGreaterThan(before);
	await removeProbe(probe);
	expect(await documentPointerListeners(page)).toBe(before);
	await expectCollected(page);
});

test('a form control list closes when its editor is removed', async ({ page }) => {
	await openFixture(page, 'vanilla', 'excel-shapes.xlsx');
	const dropdown = editor(page).locator(
		'[part="grid"] .xg-formcontrol[data-control-type="dropdown"]',
	);
	const items = dropdown.locator('[data-control-part="item"]');
	const before = await documentPointerListeners(page);
	await dropdown.locator('[data-control-part="drop"]').first().click();
	await expect(items.first()).toBeVisible();
	expect(await documentPointerListeners(page)).toBeGreaterThan(before);
	await reconnect(page);
	expect(await documentPointerListeners(page)).toBe(before);
	await expect(items).toHaveCount(0);
});

test('removing an editor leaves the popover of another editor open', async ({ page }) => {
	await newWorkbook(page);
	const second = await addEditor(page, 'second', 'right:0;bottom:0;width:640px;height:420px');
	await second
		.locator('[part="ribbon"] .ribbon-panel:not([hidden]) [aria-haspopup]:visible')
		.first()
		.click();
	await expect(second.locator('.ribbon-popover')).toHaveCount(1);
	await editor(page)
		.first()
		.evaluate((element) => element.remove());
	await expect(second.locator('.ribbon-popover')).toHaveCount(1);
});

test('an editor removed while sharing a room it opened leaves it and is collected', async ({
	page,
}) => {
	await newWorkbook(page);
	const first = await addEditor(page, 'probe', 'left:0;top:0;width:640px;height:420px');
	const second = await addEditor(page, 'second', 'right:0;bottom:0;width:640px;height:420px');
	const roomId = `lifecycle-${Date.now()}`;
	await first.evaluate((element, room) => {
		(element as Editor).collaboration = { roomId: room, user: { name: 'Ada' } };
	}, roomId);
	await expect
		.poll(() => first.evaluate((element) => (element as Editor).collaborationState.active))
		.toBe(true);
	await second.evaluate((element, room) => {
		(element as Editor).collaboration = { roomId: room, intent: 'join', user: { name: 'Bob' } };
	}, roomId);
	const people = () =>
		second.evaluate((element) => (element as Editor).collaborationState.people.length);
	await expect.poll(people).toBe(2);
	await removeProbe(first);
	await expect.poll(people).toBe(1);
	await expectCollected(page);
});
