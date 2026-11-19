import { expect, test, type Page } from '@playwright/test';
import { openSample } from './helpers';

/** Whether keyboard focus is inside the open `office-ui-dialog`, through shadow roots and slots. */
const focusInDialog = (page: Page) =>
	page.evaluate(() => {
		let node: Node | null = document.activeElement;
		while ((node as Element | null)?.shadowRoot?.activeElement)
			node = (node as Element).shadowRoot!.activeElement;
		while (node) {
			if ((node as Element).localName === 'office-ui-dialog') return true;
			node = (node as Element).assignedSlot ?? node.parentNode ?? (node as ShadowRoot).host ?? null;
		}
		return false;
	});

/** A bare dialog on the page, so nothing but the element decides where focus goes. */
const mountBareDialog = (page: Page) =>
	page.evaluate(() => {
		const opener = document.body.appendChild(document.createElement('button'));
		opener.id = 'dialog-opener';
		opener.textContent = 'Open';
		const dialog = document.createElement('office-ui-dialog') as HTMLElement & { open: boolean };
		dialog.id = 'bare-dialog';
		dialog.setAttribute('heading', 'Bare');
		dialog.innerHTML =
			'<input aria-label="First" /><select aria-label="Second"><option>A</option></select>' +
			'<button slot="footer">Cancel</button><button slot="footer">OK</button>';
		document.body.append(dialog);
		opener.addEventListener('click', () => (dialog.open = true));
		opener.focus();
	});

test.describe('office-ui-dialog keyboard contract (real Tab order)', () => {
	test('Tab and Shift+Tab stay inside, Escape closes from any stop, focus returns', async ({
		page,
	}) => {
		await openSample(page);
		await mountBareDialog(page);
		await page.keyboard.press('Enter');
		const dialog = page.locator('#bare-dialog');
		await expect(dialog).toHaveAttribute('open', '');
		await expect(dialog.getByRole('textbox', { name: 'First' })).toBeFocused();

		// close button, two fields and two footer buttons: five stops, wrapping both ways.
		for (const key of ['Tab', 'Shift+Tab']) {
			for (let press = 0; press < 12; press += 1) {
				await page.keyboard.press(key);
				expect(await focusInDialog(page), `${key} press ${press + 1}`).toBe(true);
			}
		}

		await page.keyboard.press('Tab');
		await page.keyboard.press('Escape');
		await expect(dialog).not.toHaveAttribute('open', '');
		await expect(page.locator('#dialog-opener')).toBeFocused();
	});

	test('a Word dialog keeps focus inside and returns it to the document, as Word does', async ({
		page,
	}) => {
		await openSample(page);
		const opener = page.locator('docx-editor').getByRole('button', { name: 'Font settings' });
		await opener.focus();
		await page.keyboard.press('Enter');
		const dialog = page.locator('docx-editor office-ui-dialog[open]');
		await expect(dialog).toHaveCount(1);
		for (let press = 0; press < 40; press += 1) {
			await page.keyboard.press('Tab');
			expect(await focusInDialog(page), `Tab press ${press + 1}`).toBe(true);
		}
		await page.keyboard.press('Escape');
		await expect(dialog).toHaveCount(0);
		await expect(page.locator('docx-editor .ProseMirror')).toBeFocused();
	});

	test('Escape closes even when focus has left the dialog', async ({ page }) => {
		await openSample(page);
		await mountBareDialog(page);
		await page.keyboard.press('Enter');
		const dialog = page.locator('#bare-dialog');
		await expect(dialog).toHaveAttribute('open', '');
		await page.evaluate(() => {
			let active = document.activeElement as HTMLElement | null;
			while (active?.shadowRoot?.activeElement)
				active = active.shadowRoot.activeElement as HTMLElement;
			active?.blur();
		});
		await page.keyboard.press('Escape');
		await expect(dialog).not.toHaveAttribute('open', '');
	});
});
