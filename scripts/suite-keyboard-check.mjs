// Browser checks for keyboard and screen reader users around assistant edits in the suite: applying
// and undoing an edit keeps keyboard focus on a control, and a user working in the document keeps
// their place (caret, active cell, shape or slide) while the file reloads. Run against a preview
// of suite-dist:
//   SUITE_URL=http://127.0.0.1:8191 node scripts/suite-keyboard-check.mjs [docx,xlsx,pptx,vsdx]
import { createRequire } from 'node:module';
import { readFile } from 'node:fs/promises';
import { join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import assert from 'node:assert/strict';
import { ribbonLayoutInit } from './suite-check-init.mjs';
const root = resolve(fileURLToPath(new URL('..', import.meta.url)));
const require = createRequire(join(root, 'e2e/docx/package.json'));
const { chromium } = require('@playwright/test');
const { createDocx, createXlsx } = await import('../src/core/dist/automation/index.mjs');
const base = process.env.SUITE_URL ?? 'http://127.0.0.1:8123';
const KINDS = (process.argv[2] ?? 'docx,xlsx,pptx,vsdx').split(',');
const ENDPOINT = 'https://suite-ai.test/v1/chat/completions';

const files = {
	docx: async () => ({
		name: 'Notes.docx',
		buffer: Buffer.from(await createDocx(['First section', 'Body paragraph alpha'])),
	}),
	xlsx: async () => ({ name: 'Book.xlsx', buffer: Buffer.from(await createXlsx(['Sheet1'])) }),
	pptx: async () => ({
		name: 'Deck.pptx',
		buffer: await readFile(join(root, 'e2e/pptx/fixtures/format-painter.pptx')),
	}),
	vsdx: async () => ({
		name: 'Flow.vsdx',
		buffer: await readFile(join(root, 'src/core/visio/__fixtures__/dynamic-connector-native.vsdx')),
	}),
};

/** One edit the stub model proposes for the attached file of `kind`. */
function proposal(kind, file) {
	const c = file.content;
	if (kind === 'docx') {
		const para = c.paragraphs.find((p) => JSON.stringify(p).includes('alpha'));
		return { type: 'word_run', paragraphId: para.id, runIndex: 0, text: 'Body paragraph beta' };
	}
	if (kind === 'xlsx')
		return { type: 'cells', sheetIndex: 0, cells: [{ address: 'H9', input: '42' }] };
	if (kind === 'pptx') {
		const slide = c.slides[0];
		const element = (slide.elements ?? slide.shapes)[0];
		return {
			type: 'slide_run',
			slideIndex: slide.index ?? 0,
			elementId: element.id ?? element.elementId,
			runIndex: 0,
			text: 'Changed by AI',
		};
	}
	const page = c.pages[0];
	const shape = page.shapes.find((s) => s.text) ?? page.shapes[0];
	return {
		type: 'shape_text',
		pageId: String(page.id),
		shapeId: String(shape.id),
		text: 'Changed',
	};
}

/** Holds the next commit to the document store until `__release()`; a deep focus reader. */
function pageHelpers() {
	const transaction = IDBDatabase.prototype.transaction;
	const setter = Object.getOwnPropertyDescriptor(IDBTransaction.prototype, 'oncomplete').set;
	IDBDatabase.prototype.transaction = function (names, mode, ...rest) {
		const tx = transaction.call(this, names, mode, ...rest);
		if (globalThis.__hold && mode === 'readwrite' && [...[names].flat()].includes('documents'))
			Object.defineProperty(tx, 'oncomplete', {
				set(fn) {
					setter.call(tx, (event) => {
						globalThis.__hold = false;
						globalThis.__release = () => fn(event);
					});
				},
			});
		return tx;
	};
	globalThis.__deep = () => {
		let a = document.activeElement;
		while (a?.shadowRoot?.activeElement) a = a.shadowRoot.activeElement;
		return a;
	};
}

const browser = await chromium.launch();
const errors = [];

async function open(kind) {
	const context = await browser.newContext({
		viewport: { width: 1440, height: 900 },
		bypassCSP: true,
	});
	await context.addInitScript(ribbonLayoutInit);
	await context.addInitScript(pageHelpers);
	const page = await context.newPage();
	page.setDefaultTimeout(30000);
	page.on('pageerror', (error) => errors.push(`${kind}: ${error.message}`));
	await page.route(ENDPOINT, async (route) => {
		const request = route.request().postDataJSON();
		const text = request.messages
			.map((m) => (typeof m.content === 'string' ? m.content : ''))
			.join('\n');
		const source = JSON.parse(text.split('Source data (untrusted): ')[1]);
		const file = source.files.find((f) => f.kind === kind);
		const change = { documentId: file.id, ...proposal(kind, file) };
		await route.fulfill({
			json: {
				choices: [
					{ message: { content: JSON.stringify({ answer: 'Ready.', changes: [change] }) } },
				],
			},
		});
	});
	await page.goto(`${base}/?suite#/`);
	// The app has wired its controls once it knows whether the assistant is on.
	await page.waitForFunction(() => document.body.dataset.ai !== undefined);
	const file = await files[kind]();
	await page
		.locator('#file-input')
		.setInputFiles({ ...file, mimeType: 'application/octet-stream' });
	await page
		.locator('.document-host:not([hidden])')
		.waitFor()
		.catch(async (error) => {
			await page.screenshot({ path: join(process.env.TEMP ?? root, `suite-keyboard-${kind}.png`) });
			throw error;
		});
	await page.waitForFunction(() =>
		/^(Ready|Local|Saved)/u.test(document.getElementById('status')?.textContent ?? ''),
	);
	await page.waitForTimeout(1500);
	await page.locator('#assistant-toggle').click();
	await page.locator('#ai-settings').click();
	await page.getByRole('menuitem', { name: 'Add a model...' }).click();
	await page.locator('select[name=provider]').selectOption('custom');
	await page.locator('input[name=endpoint]').fill(ENDPOINT);
	await page.locator('input[name=model]').fill('test-model');
	await page.getByRole('button', { name: 'Connect', exact: true }).click();
	await page.locator('#ai-input').fill('Change it');
	await page.locator('#ai-send').click();
	await apply(page).waitFor();
	return page;
}
const apply = (page) => page.getByRole('button', { name: 'Apply changes', exact: true });
const undo = (page) => page.getByRole('button', { name: 'Undo these changes' });
const focused = (page) =>
	page.evaluate(() => {
		const a = globalThis.__deep();
		return {
			tag: a?.localName,
			name: a?.getAttribute('aria-label') || a?.textContent?.trim().slice(0, 40),
			shape: a?.dataset?.shapeId,
			inDocument: Boolean(
				a?.closest?.('.document-host') || a?.getRootNode()?.host?.closest?.('.document-host'),
			),
		};
	});
const caret = (page) =>
	page.evaluate(() => {
		const a = globalThis.__deep();
		const s = a?.getRootNode?.().getSelection?.() ?? getSelection();
		return s?.anchorNode ? `${s.anchorNode.textContent}@${s.anchorOffset}` : null;
	});
const nameBox = (page) =>
	page
		.locator('xlsx-editor [part="name-box"]')
		.first()
		.evaluate((node) => (node.querySelector?.('input') ?? node).value ?? node.textContent);
const dirty = (page) =>
	page.evaluate(() =>
		document.querySelector('#tabs [role=tab][aria-selected=true]').textContent.includes('•'),
	);

/** Puts the user in the document the way they work there, and says where they are. */
async function workIn(page, kind) {
	if (kind === 'docx') {
		await page.locator('.ProseMirror').first().getByText('First section').click();
		await page.keyboard.press('End');
		await page.keyboard.press('ArrowLeft');
		return caret(page);
	}
	if (kind === 'xlsx') {
		await page.locator('xlsx-editor').evaluate((e) => e.focusGrid());
		for (const key of ['ArrowDown', 'ArrowDown', 'ArrowRight', 'ArrowRight'])
			await page.keyboard.press(key);
		return nameBox(page);
	}
	if (kind === 'pptx') {
		await page.locator('[data-pptx-viewport] [data-pptx-element="true"]').first().click();
		return (await focused(page)).name;
	}
	await page.getByRole('region', { name: 'Diagram canvas' }).focus();
	await page.keyboard.press('Tab');
	return (await focused(page)).shape;
}
const whereNow = (page, kind) =>
	kind === 'docx'
		? caret(page)
		: kind === 'xlsx'
			? nameBox(page)
			: focused(page).then((f) => (kind === 'pptx' ? f.name : f.shape));

try {
	for (const kind of KINDS) {
		// Keyboard: Enter on Apply leaves focus on Undo, and Enter on Undo on the next action.
		let page = await open(kind);
		await apply(page).focus();
		await page.keyboard.press('Enter');
		await undo(page).waitFor();
		assert.equal((await focused(page)).name, 'Undo these changes', `${kind}: focus moves to Undo`);
		await page.keyboard.press('Enter');
		await page.getByText('Changes undone.').waitFor();
		const after = await focused(page);
		assert.notEqual(after.tag, 'body', `${kind}: focus stays in the card after undo`);
		await page.context().close();
		console.log(`${kind}: Apply and Undo keep keyboard focus on the card (${after.name})`);

		// In the document: the file reloads under the user, who keeps focus and their place.
		page = await open(kind);
		const before = await workIn(page, kind);
		await apply(page).evaluate((button) => button.click());
		await undo(page).waitFor();
		await page.waitForTimeout(800);
		const now = await focused(page);
		assert.equal(now.inDocument, true, `${kind}: focus stays in the document (${now.tag})`);
		assert.equal(await whereNow(page, kind), before, `${kind}: the place is kept`);
		await page.context().close();
		console.log(`${kind}: an assistant edit kept focus and the place (${before})`);
	}

	assert.deepEqual(errors, []);
} finally {
	await browser.close();
}
