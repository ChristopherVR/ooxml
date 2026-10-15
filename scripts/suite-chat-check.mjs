// Browser checks for the assistant's workspace features: the assistant stays off without an AI
// configuration, files upload from the Attach menu, a tab dropped on the chat is attached, tabs
// reorder, and a shared chat syncs between two pages. Run against a preview of suite-dist:
//   SUITE_URL=http://localhost:8123 node scripts/suite-chat-check.mjs
import { createRequire } from 'node:module';
import { join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import assert from 'node:assert/strict';
const root = resolve(fileURLToPath(new URL('..', import.meta.url)));
const require = createRequire(join(root, 'e2e/docx/package.json'));
const { chromium } = require('@playwright/test');
const { createDocx } = await import('../src/core/dist/automation/index.mjs');
const base = process.env.SUITE_URL ?? 'http://localhost:8123';
const browser = await chromium.launch({ headless: true });
const errors = [];
const watch = (page) => page.on('pageerror', (error) => errors.push(error.message));
try {
	// Without a configuration the suite has no assistant and opens on its files.
	const plain = await browser.newContext({ viewport: { width: 1280, height: 900 } });
	await plain.route('**/ai-config.json', (route) => route.fulfill({ status: 404, body: '' }));
	await plain.addInitScript(() => (globalThis.ooxmlSkipStart = true));
	const off = await plain.newPage();
	watch(off);
	await off.goto(base);
	await off.locator('[data-create="docx"]').waitFor();
	await off.waitForFunction(() => document.body.dataset.ai === 'off');
	for (const selector of ['#home-ask', '#assistant-toggle', '#rail [data-chat]'])
		assert.equal(await off.locator(selector).isVisible(), false, `${selector} is hidden`);
	await off.goto(`${base}/#/chat`);
	await off.waitForFunction(() => location.hash === '#/');
	await plain.close();
	console.log('No configuration: no assistant, the suite opens on its files');

	const context = await browser.newContext({ viewport: { width: 1280, height: 900 } });
	await context.addInitScript(() => (globalThis.ooxmlSkipStart = true));
	const page = await context.newPage();
	watch(page);
	await page.goto(base);
	await page.locator('#chat-input').waitFor();

	// Upload from the Attach menu: an Office file joins the library, a text file becomes a note.
	const brief = Buffer.from(await createDocx(['Launch brief', 'Ships in November.']));
	const chooser = page.waitForEvent('filechooser');
	await page.locator('#chat-context').click();
	await page.getByRole('menuitem', { name: 'Upload from this device...' }).click();
	await (
		await chooser
	).setFiles([
		{ name: 'Brief.docx', mimeType: 'application/octet-stream', buffer: brief },
		{
			name: 'notes.txt',
			mimeType: 'text/plain',
			buffer: Buffer.from('Budget is 4200.\nOwner: Ada'),
		},
	]);
	await page.locator('#chat-scope .chat-chip', { hasText: 'Brief.docx' }).waitFor();
	await page.locator('#chat-scope .chat-chip', { hasText: 'notes.txt' }).waitFor();
	await page.locator('#chat-input').fill('Describe the attached files');
	await page.locator('#chat-send').click();
	await page.locator('.chat-msg.from-assistant', { hasText: 'notes.txt' }).waitFor();
	await page.locator('.chat-msg.from-assistant', { hasText: 'Brief.docx' }).waitFor();
	console.log('Uploaded a Word file and a text note from the Attach menu');

	// Open two documents, reorder their tabs, then drop one on the assistant to attach it.
	for (const kind of ['xlsx', 'pptx']) {
		await page.locator('#rail [data-home]').click();
		await page.locator(`[data-create="${kind}"]`).first().click();
		await page
			.locator('.file-tab')
			.nth(kind === 'xlsx' ? 0 : 1)
			.waitFor();
	}
	const names = () =>
		page.locator('.file-tab [role=tab]').evaluateAll((tabs) => tabs.map((t) => t.textContent));
	const before = await names();
	await page.locator('.file-tab').last().locator('[role=tab]').focus();
	await page.keyboard.press('Control+Shift+ArrowLeft');
	assert.deepEqual(await names(), [...before].reverse(), 'Ctrl+Shift+Left moves the tab');
	await page
		.locator('.file-tab')
		.last()
		.dragTo(page.locator('.file-tab').first(), {
			targetPosition: { x: 4, y: 8 },
		});
	assert.deepEqual(await names(), before, 'dragging puts the tab back');
	console.log('Tabs reorder by keyboard and by dragging');

	await page.locator('#assistant-toggle').click();
	await page.locator('#ai-input').waitFor();
	for (const chip of await page.locator('#ai-scope [data-detach]').all()) await chip.click();
	const dropped = await page.locator('.file-tab').first().getAttribute('data-tab');
	await page.locator('.file-tab').first().dragTo(page.locator('#assistant-chat .chat-composer'));
	await page.locator('#ai-scope .chat-chip').first().waitFor();
	const attachedName = await page.locator('.file-tab').first().locator('[role=tab]').textContent();
	assert.match(
		attachedName,
		new RegExp((await page.locator('#ai-scope .chat-chip').first().textContent()).slice(0, 8)),
	);
	assert.ok(dropped);
	console.log('Dropping a tab on the assistant attaches it');

	// Share a chat, then join it from a second page and watch a request arrive.
	await page.goto(`${base}/#/chat`);
	await page.locator('[data-chat-new]').click();
	await page.locator('#chat-input').fill('Create a presentation about our product launch');
	await page.locator('#chat-send').click();
	await page.locator('.chat-msg.from-assistant .ai-create-card').waitFor();
	await page.locator('[data-chat-share]').click();
	const link = await page.locator('input[name=link]').inputValue();
	assert.match(link, /#\/chat\/join\/chat-[0-9a-f]{32}$/u);
	await page.locator('#dialog-close').click();
	await page.locator('#chat-people').waitFor();

	const guest = await context.newPage();
	watch(guest);
	await guest.goto(link);
	await guest.locator('.chat-msg.from-user', { hasText: 'product launch' }).waitFor();
	await guest.locator('.chat-msg.from-assistant .ai-create-card').waitFor();
	await page.locator('#chat-people', { hasText: 'other' }).waitFor();
	await guest.locator('#chat-input').fill('Make a budget workbook for the launch');
	await guest.locator('#chat-send').click();
	await page.locator('.chat-msg.from-user', { hasText: 'budget workbook' }).waitFor();
	await page.locator('.chat-msg.from-assistant', { hasText: 'budget.xlsx' }).waitFor();
	console.log('A shared chat synced requests and replies between two pages');
	assert.deepEqual(errors, []);
	await context.close();
} finally {
	await browser.close();
}
