import { test, expect, type Locator } from '@playwright/test';

for (const guest of ['vue', 'solid']) {
	test(`local coauthor demo relays accessible presence to ${guest}`, async ({ page }) => {
		const errors: string[] = [];
		page.on('pageerror', (error) => errors.push(error.message));
		await page.goto(`/collaboration.html?framework=vanilla&guest=${guest}`);
		const peerA = page.locator('#peer-a docx-editor');
		const peerB = page.locator('#peer-b docx-editor');
		const aliceCursor = peerB.getByRole('img', { name: "Ada's cursor" });
		const graceCursor = peerA.getByRole('img', { name: "Grace's cursor" });
		await expect(aliceCursor).toBeVisible();
		await expect(graceCursor).toBeVisible();
		await page.getByLabel('Editor B cursor color').selectOption('#a21caf');
		await page.getByLabel('Name', { exact: true }).nth(1).fill('Grace Hopper');
		const updatedCursor = peerA.getByRole('img', { name: "Grace Hopper's cursor" });
		await expect(updatedCursor).toBeVisible();
		await expect(updatedCursor).toHaveCSS('border-left-color', 'rgb(162, 28, 175)');
		await expect(peerA.getByRole('img', { name: "Grace's cursor" })).toHaveCount(0);

		await peerB.evaluate((element) =>
			(element as unknown as { leavePresence: () => unknown }).leavePresence(),
		);
		await expect(peerA.getByRole('img', { name: "Grace Hopper's cursor" })).toHaveCount(0);
		await expect(page.getByRole('status').filter({ hasText: 'Synced' })).toBeVisible();
		await expect(errors).toEqual([]);
	});
}

for (const guest of ['vue', 'solid']) {
	test(`Yjs relative selections and profile updates reach ${guest}`, async ({ page }) => {
		await page.addInitScript(() => {
			(window as unknown as { wordErrors: string[] }).wordErrors = [];
			document.addEventListener('document-error', (event) =>
				(window as unknown as { wordErrors: string[] }).wordErrors.push(
					(event as CustomEvent<Error>).detail.message,
				),
			);
		});
		const errors: string[] = [];
		page.on('pageerror', (error) => errors.push(error.message));
		await page.goto(`/collaboration.html?framework=vanilla&guest=${guest}&mode=yjs`);
		const a = page.locator('#peer-a docx-editor');
		const b = page.locator('#peer-b docx-editor');
		await expect(b.locator('.ProseMirror')).toContainText('Shared document');
		await b.locator('.ProseMirror').click();
		await page.keyboard.press('Control+Home');
		for (let i = 0; i < 5; i++) await page.keyboard.press('Shift+ArrowRight');
		await expect(a.getByRole('img', { name: "Grace's cursor" })).toBeVisible();
		await expect(a.locator('.dve-peer-selection')).toHaveText('Share');
		await a.evaluate((element) => {
			const editor = element as unknown as { view: import('prosemirror-view').EditorView };
			editor.view.dispatch(editor.view.state.tr.insertText('Prefix ', 1));
		});
		await expect.poll(() => bodyText(a)).toBe('Prefix Shared document');
		await expect.poll(() => bodyText(b)).toBe('Prefix Shared document');
		await expect(a.locator('.dve-peer-selection')).toHaveText('Share');
		await page.getByLabel('Editor B cursor color').selectOption('#a21caf');
		await page.getByLabel('Name', { exact: true }).nth(1).fill('Grace Hopper');
		await b.locator('.ProseMirror').click();
		const cursor = a.getByRole('img', { name: "Grace Hopper's cursor" });
		await expect(cursor).toBeVisible();
		await expect(cursor).toHaveCSS('border-left-color', 'rgb(162, 28, 175)');
		await expect.poll(() => bodyText(b)).toBe('Prefix Shared document');
		await page.getByRole('button', { name: 'Reconnect providers', exact: true }).click();
		await page.getByRole('button', { name: 'Resync providers', exact: true }).click();
		await expect.poll(() => bodyText(b)).toBe('Prefix Shared document');
		await b.evaluate((element) =>
			(element as unknown as { leavePresence(): unknown }).leavePresence(),
		);
		await expect(cursor).toHaveCount(0);
		expect(
			await page.evaluate(() => (window as unknown as { wordErrors: string[] }).wordErrors),
		).toEqual([]);
		expect(errors).toEqual([]);
	});
}

function bodyText(peer: Locator): Promise<string> {
	return peer.evaluate(
		(element) =>
			(element as unknown as { view: import('prosemirror-view').EditorView }).view.state.doc
				.textContent,
	);
}
