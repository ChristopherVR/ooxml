// Browser check for the suite shell on phones and tablets: the app-rail drawer covers editor
// overlays, tabs keep their width and clip cleanly, the command bar's assistant button is a plain
// icon button, and the chat page is a list or a conversation (never both). Runs at four sizes in
// both colour schemes against a preview of suite-dist:
//   SUITE_URL=http://localhost:8123 node scripts/suite-mobile-check.mjs [screenshot-dir]
import { createRequire } from 'node:module';
import { join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import assert from 'node:assert/strict';
import { mkdir } from 'node:fs/promises';
const root = resolve(fileURLToPath(new URL('..', import.meta.url)));
const require = createRequire(join(root, 'e2e/docx/package.json'));
const { chromium } = require('@playwright/test');
const base = process.env.SUITE_URL ?? 'http://localhost:8123';
const shots = process.argv[2] ? resolve(process.argv[2]) : '';
if (shots) await mkdir(shots, { recursive: true });
const sizes = [
	[375, 812],
	[412, 915],
	[812, 375],
	[768, 1024],
];
const browser = await chromium.launch({ headless: true });
const errors = [];
try {
	for (const scheme of ['light', 'dark'])
		for (const [width, height] of sizes) {
			const label = `${scheme} ${width}x${height}`;
			const context = await browser.newContext({
				viewport: { width, height },
				hasTouch: true,
				isMobile: true,
				colorScheme: scheme,
			});
			await context.addInitScript(() => (globalThis.ooxmlSkipStart = true));
			// The suite opens dark; choose the mode through its own appearance preference.
			await context.addInitScript(
				(mode) =>
					localStorage.setItem('ooxml-suite-theme', JSON.stringify({ mode, accent: 'blue' })),
				scheme,
			);
			const page = await context.newPage();
			page.on('pageerror', (error) => errors.push(`${label}: ${error.message}`));
			const shot = (name) =>
				shots
					? page.screenshot({ path: join(shots, `${scheme}-${width}x${height}-${name}.png`) })
					: 0;
			const noOverflow = async (what) =>
				assert.equal(
					await page.evaluate(() => document.documentElement.scrollWidth > innerWidth),
					false,
					`${label}: ${what} scrolls the page sideways`,
				);
			await page.goto(base);
			await page.locator('[data-create]').first().waitFor({ state: 'attached' });
			assert.equal(
				await page.evaluate(() => document.documentElement.dataset.theme),
				scheme,
				`${label}: theme applied`,
			);
			await noOverflow('home');
			await shot('home');
			for (const kind of ['docx', 'xlsx', 'pptx']) {
				await page.evaluate((k) => document.querySelector(`[data-create="${k}"]`).click(), kind);
				await page.waitForTimeout(1200);
			}
			await noOverflow('three editors');
			await shot('editor');

			// Tabs keep their natural width: a shrunk tab clips its own label.
			const clipped = await page.evaluate(() =>
				[...document.querySelectorAll('#tabs > button, #tabs .file-tab > button:first-child')]
					.filter(
						(b) => b.offsetParent && b.getBoundingClientRect().width < 30 && b.textContent.trim(),
					)
					.map((b) => b.textContent.trim()),
			);
			assert.deepEqual(clipped, [], `${label}: tabs shrunk below their label`);

			// The collapse toggle (phone landscape) never covers the editor.
			const overlap = await page.evaluate(() => {
				const t = document.querySelector('#chrome-toggle');
				const e = document.querySelector('#editors');
				if (!t || !e || getComputedStyle(t).display === 'none') return false;
				const a = t.getBoundingClientRect();
				const b = e.getBoundingClientRect();
				return a.left < b.right && a.right > b.left && a.top < b.bottom && a.bottom > b.top;
			});
			assert.equal(overlap, false, `${label}: the chrome toggle covers the editor`);

			// The command bar's assistant button is not an outlined box on a narrow screen.
			if (width <= 700) {
				const border = await page.evaluate(() => {
					const b = document.querySelector('.document-commandbar .ask-ai');
					return b && b.offsetParent ? getComputedStyle(b).borderTopColor : 'rgba(0, 0, 0, 0)';
				});
				assert.match(border, /rgba\(\d+, \d+, \d+, 0\)|transparent/u, `${label}: ask-ai border`);
			}

			// The drawer covers the editor, including overlays that sit fixed at the bottom.
			if (width <= 640) {
				await page.locator('#launcher-toggle').click();
				await page.waitForTimeout(500);
				await shot('drawer');
				const top = await page.evaluate(() => {
					const r = document.querySelector('.workspace-sidebar').getBoundingClientRect();
					const probes = [r.bottom - 10, r.top + 10, r.top + r.height / 2];
					return probes.map((y) => {
						const e = document.elementFromPoint(r.left + 20, y);
						return !!e?.closest('.workspace-sidebar');
					});
				});
				assert.deepEqual(top, [true, true, true], `${label}: the drawer is covered by content`);
				await page.locator('#drawer-scrim').click({ position: { x: width - 10, y: 200 } });
			}

			// Chat: a list or a conversation, never a list squeezed beside it.
			await page.goto(`${base}/#/chat`);
			await page.reload();
			await page.locator('.chat-page-chat textarea').first().waitFor();
			await shot('chat');
			if (width <= 640) {
				const closed = await page.evaluate(
					() => getComputedStyle(document.querySelector('.chat-history')).visibility,
				);
				assert.equal(closed, 'hidden', `${label}: the chat list is hidden on the conversation`);
				await page.locator('.chat-history-toggle').click();
				await page.waitForTimeout(400);
				const list = await page.locator('.chat-history').boundingBox();
				assert.ok(list.width >= width - 1, `${label}: the chat list is full width`);
				await shot('chat-list');
				await page.locator('.chat-history-back').click();
				await page.waitForTimeout(400);
				assert.equal(
					await page.evaluate(
						() => getComputedStyle(document.querySelector('.chat-history')).visibility,
					),
					'hidden',
					`${label}: back returns to the conversation`,
				);
			}
			await noOverflow('chat');
			await context.close();
			console.log(`${label}: shell ok`);
		}
	assert.deepEqual(errors, []);
} finally {
	await browser.close();
}
