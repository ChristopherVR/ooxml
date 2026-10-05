import { expect, test } from '@playwright/test';

// A `?room=` session is only a name on a BroadcastChannel, so windows built from different
// framework adapters must meet in it: the host (`sample=1`) opens the sample and the guest joins by
// name, receives the document and then exchanges edits and cursors.
for (const [host, guest] of [
	['react', 'vue'],
	['vanilla', 'angular'],
	['svelte', 'solid'],
] as const) {
	test(`a ${host} window and a ${guest} window share one document`, async ({ context }) => {
		const room = `e2e-${host}-${guest}-${Date.now().toString(36)}`;
		const a = await context.newPage();
		const b = await context.newPage();
		await a.goto(`/?framework=${host}&sample=1&room=${room}&name=Ada`);
		const textA = a.locator('docx-editor .ProseMirror');
		await expect(textA).toContainText('Document title');
		await b.goto(`/?framework=${guest}&room=${room}&name=Grace`);
		const textB = b.locator('docx-editor .ProseMirror');
		// The guest starts blank and receives the host's document.
		await expect(textB).toContainText('Document title');

		await textA.click();
		await a.keyboard.press('Control+End');
		await a.keyboard.type(` from-${host}`);
		await expect(textB).toContainText(`from-${host}`);

		await textB.click();
		await b.keyboard.press('Control+End');
		await b.keyboard.type(` from-${guest}`);
		await expect(textA).toContainText(`from-${guest}`);
	});
}
