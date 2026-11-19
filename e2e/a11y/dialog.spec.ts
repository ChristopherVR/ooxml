/**
 * The shared dialogs' Tab order, in a real browser: `office-ui-dialog` walks its close button,
 * fields, the controls inside a composite element and its footer in order and wraps both ways,
 * takes Tab back when focus has ended up behind it, and `office-ui-options-dialog` lets Tab walk a
 * pane a host slots in. Mounted bare on the suite page (after Word has registered the elements) so
 * nothing but the elements decides where focus goes.
 */
import type { Page } from '@playwright/test';
import { expect, openSurface, test } from './support/suite';

test.use({ viewport: { width: 1440, height: 900 } });

/** The focused element's label, through shadow roots, or `outside` when it is not in `#bare`. */
const stop = (page: Page) =>
	page.evaluate(() => {
		const active = globalThis.__a11y.deep();
		if (!active) return 'body';
		if (!globalThis.__a11y.focusWithin('#bare')) return `outside: ${active.localName}`;
		return active.getAttribute('aria-label') || active.textContent?.trim() || active.localName;
	});

async function walk(page: Page, key: 'Tab' | 'Shift+Tab', presses: number): Promise<string[]> {
	const stops: string[] = [];
	for (let press = 0; press < presses; press += 1) {
		await page.keyboard.press(key);
		stops.push(await stop(page));
	}
	return stops;
}

test.beforeEach(async ({ page }) => {
	await openSurface(page, 'docx');
});

test('office-ui-dialog walks its stops in order and wraps both ways', async ({ page }) => {
	await page.evaluate(() => {
		const opener = document.body.appendChild(document.createElement('button'));
		opener.id = 'opener';
		opener.textContent = 'Open';
		const dialog = document.createElement('office-ui-dialog') as HTMLElement & { open: boolean };
		dialog.id = 'bare';
		dialog.setAttribute('heading', 'Bare');
		const picker = document.createElement('div');
		picker.attachShadow({ mode: 'open' }).innerHTML =
			'<button aria-label="Picker first">A</button><button aria-label="Picker last">B</button>';
		const field = document.createElement('input');
		field.setAttribute('aria-label', 'Field');
		const cancel = Object.assign(document.createElement('button'), {
			slot: 'footer',
			textContent: 'Cancel',
		});
		const ok = Object.assign(document.createElement('button'), {
			slot: 'footer',
			textContent: 'OK',
		});
		dialog.append(field, picker, cancel, ok);
		document.body.append(dialog);
		opener.addEventListener('click', () => (dialog.open = true));
		opener.focus();
	});
	await page.keyboard.press('Enter');
	await expect.poll(() => stop(page)).toBe('Field');
	const close = (await walk(page, 'Shift+Tab', 1))[0];
	expect(close).not.toMatch(/^outside|^body/u);
	const order = [close, 'Field', 'Picker first', 'Picker last', 'Cancel', 'OK'];
	expect(await walk(page, 'Tab', 12), 'Tab walks every stop in order and wraps').toEqual([
		...order.slice(1),
		...order,
		order[0],
	]);
	expect(await walk(page, 'Shift+Tab', 6), 'Shift+Tab walks them back').toEqual(
		[...order].reverse(),
	);

	// Focus that has ended up behind the dialog comes back with the next Tab.
	await page.locator('#opener').focus();
	await page.keyboard.press('Tab');
	expect(await stop(page)).not.toMatch(/^outside|^body/u);
	await page.locator('#opener').focus();
	await page.keyboard.press('Escape');
	await expect(page.locator('#bare')).not.toHaveAttribute('open', '');
});

test('office-ui-options-dialog lets Tab walk a slotted pane and keeps it inside', async ({
	page,
}) => {
	await page.evaluate(async () => {
		const dialog = document.createElement('office-ui-options-dialog') as HTMLElement & {
			categories: unknown[];
			show(): void;
			updateComplete: Promise<unknown>;
		};
		dialog.id = 'bare';
		dialog.categories = [
			{ id: 'ribbon', label: 'Customize Ribbon', sections: [], custom: true },
			{ id: 'other', label: 'Other', sections: [] },
		];
		for (const label of ['Pane first', 'Pane last']) {
			const field = document.createElement('input');
			field.slot = 'pane-ribbon';
			field.setAttribute('aria-label', label);
			dialog.append(field);
		}
		document.body.append(dialog);
		dialog.show();
		await dialog.updateComplete;
	});
	const stops = await walk(page, 'Tab', 20);
	expect(
		stops.filter((label) => /^outside|^body/u.test(label)),
		'Tab stays inside',
	).toEqual([]);
	expect(stops, 'Tab reaches both fields of the slotted pane, in order').toEqual(
		expect.arrayContaining(['Pane first', 'Pane last']),
	);
	expect(stops.indexOf('Pane last')).toBe(stops.indexOf('Pane first') + 1);
	const back = await walk(page, 'Shift+Tab', 20);
	expect(
		back.filter((label) => /^outside|^body/u.test(label)),
		'Shift+Tab stays inside',
	).toEqual([]);
});
