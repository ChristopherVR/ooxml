/** Check the built Pages demos, including each route's framework default. Serve pages-dist first. */
import assert from 'node:assert/strict';
import { createRequire } from 'node:module';
import { createVsdxFixture } from '../e2e/visio/fixture.mjs';

const require = createRequire(new URL('../e2e/docx/package.json', import.meta.url));
const { chromium, expect } = require('@playwright/test');
const base = (process.env.DEMO_URL ?? 'http://localhost:8123/ooxml').replace(/\/$/, '');
const frameworks = ['react', 'vue', 'angular', 'svelte', 'solid', 'vanilla'];
const browser = await chromium.launch({ headless: true });
const failures = [];

async function theme(page, editor, canvas, paper) {
	const node = await editor.elementHandle();
	const content = await paper.textContent();
	const paints = [];
	for (const mode of ['light', 'dark', 'light']) {
		await page.evaluate((value) => {
			localStorage.setItem('vitepress-theme-appearance', value);
			window.dispatchEvent(
				new StorageEvent('storage', {
					key: 'vitepress-theme-appearance',
					newValue: value,
				}),
			);
		}, mode);
		await expect(page.locator('html')).toHaveAttribute('data-theme', mode);
		paints.push(await canvas.evaluate((element) => getComputedStyle(element).backgroundColor));
		await expect.poll(() => paper.textContent()).toBe(content);
		assert(await editor.evaluate((element, original) => element === original, node));
	}
	assert.notEqual(paints[0], paints[1], 'Theme must change the editor chrome');
	assert.equal(paints[0], paints[2], 'Returning to light must restore its colors');
}

async function word(page, framework) {
	await expect(page.locator('#build-stamp')).toContainText(`demo · ${framework}`);
	await page.locator('#sample').click();
	const editor = page.locator('docx-editor');
	const paper = editor.locator('.ProseMirror');
	await expect(paper).toContainText('Document title');
	await paper.click();
	await page.keyboard.press('Control+End');
	await page.keyboard.type(' Framework regression check');
	await expect(paper).toContainText('Framework regression check');
	await theme(page, editor, editor.locator('.dve-canvas'), paper);
	const state = await page.evaluate(async () => {
		const handle = document.querySelector('#editor').docxEditorHandle;
		const dirty = handle.dirty;
		const saved = await handle.save();
		await handle.load(await saved.arrayBuffer());
		handle.markClean();
		return { dirty, clean: !handle.dirty };
	});
	assert.deepEqual(state, { dirty: true, clean: true });
	await expect(paper).toContainText('Framework regression check');
}

async function excel(page, framework) {
	await expect(page.locator('#build-stamp')).toContainText(`demo · ${framework}`);
	await page.locator('#sample').click();
	const editor = page.locator('xlsx-editor');
	await expect(editor.locator('[part="sheet-tabs"]')).toContainText('Sales');
	await editor.evaluate((element) => {
		element.select('B20');
		const props = element.getToolbarProps();
		props.onFormulaInput('=1+2');
		if (!props.commitEdit()) throw new Error('Formula commit rejected');
	});
	const formula = editor.locator('[part="formula-bar"]');
	await expect.poll(() => editor.evaluate((e) => e.getToolbarProps().formula)).toBe('=1+2');
	await theme(page, editor, editor.locator('.xg-hbox').first(), formula);
	assert.equal(
		await editor.evaluate(async (element) => {
			const bytes = await element.saveBytes();
			await element.load(bytes);
			element.select('B20');
			return element.getToolbarProps().formula;
		}),
		'=1+2',
	);
}

async function visio(page) {
	const editor = page.locator('visio-viewer');
	await expect(editor.locator('svg.paper')).toHaveAttribute('aria-label', 'Release workflow');
	await page.locator('#file').setInputFiles({
		name: 'framework-check.vsdx',
		mimeType: 'application/octet-stream',
		buffer: await createVsdxFixture('Before framework edit'),
	});
	await expect(page.locator('#file-name')).toHaveText('framework-check.vsdx');
	await expect(editor.locator('svg.paper')).toContainText('Before framework edit');
	await editor.locator('svg.paper [data-shape-id="1"]').dblclick();
	const input = editor.locator('#edit-text');
	await input.fill('After framework edit');
	await input.press('Escape');
	await expect(editor.locator('svg.paper')).toContainText('After framework edit');
	await theme(page, editor, editor.locator('.viewport'), editor.locator('svg.paper'));
	await expect(page.locator('#edit-label')).toHaveText('EDITED COPY');
}

async function powerpoint(page) {
	await expect(page.locator('[data-pptx-viewport] [data-pptx-element]').first()).toBeVisible();
	const viewer = page.locator('[data-pptx-theme-root]').first();
	const slides = page.locator('[aria-label="Go to slide 2"]').first();
	await slides.click();
	await expect(page.locator('[data-pptx-viewport]')).toContainText('Agenda');
	await page.locator('[aria-label="Go to slide 1"]').first().click();
	await expect(page.locator('[data-pptx-viewport]')).toContainText('Project');
	const content = await page.locator('[data-pptx-viewport]').textContent();
	// The demos explicitly supply legacy ViewerTheme inputs. Remove them to exercise Office fallback.
	await page.evaluate(() => {
		for (const element of document.querySelectorAll('[style]'))
			for (const key of Array.from(element.style))
				if (key.startsWith('--pptx-') || key.startsWith('--color-'))
					element.style.removeProperty(key);
	});
	const paints = [];
	for (const mode of ['light', 'dark', 'light']) {
		await viewer.evaluate((element, value) => {
			element.dataset.officeTheme = value;
		}, mode);
		paints.push(await viewer.evaluate((element) => getComputedStyle(element).backgroundColor));
		await expect.poll(() => page.locator('[data-pptx-viewport]').textContent()).toBe(content);
	}
	assert.notEqual(paints[0], paints[1]);
	assert.equal(paints[0], paints[2]);
	await viewer.evaluate((element) => {
		element.style.setProperty('--office-background', 'rgb(23, 45, 67)');
	});
	assert.equal(
		await viewer.evaluate((element) => getComputedStyle(element).backgroundColor),
		'rgb(23, 45, 67)',
	);
}

async function nativeBlocks(page, product, framework, route) {
	const word = product === 'docx';
	await page.goto(
		`${base}/${product}/${route}/${word ? `building-blocks.html?framework=${framework}` : '?blocks=1'}`,
	);
	if (!word) await page.locator('#blank').click();
	const toolbar = page.locator('[data-office-block="toolbar"]');
	await expect(toolbar).toBeVisible();
	const editor = page.locator(word ? 'docx-editor' : 'xlsx-editor');
	const original = await editor.elementHandle();
	const bold = page.getByRole('button', { name: 'Bold', exact: true });
	if (word) {
		await editor.locator('.ProseMirror').click();
		await page.keyboard.press('End');
	} else {
		await page.evaluate(() => window.xlsxNativeBlocks.handle.select('B20'));
		await expect(page.getByRole('textbox', { name: 'Name Box', exact: true })).toHaveCount(1);
		const formula = page.getByRole('textbox', { name: 'Formula Bar', exact: true });
		await formula.fill('=1+2');
		await formula.press('Enter');
	}
	await bold.click();
	await expect(bold).toHaveAttribute('aria-pressed', 'true');
	if (word) {
		await page.keyboard.type(' native regression');
		await expect(editor.locator('.ProseMirror strong')).toContainText('native regression');
	}
	const colors = [];
	for (const mode of ['light', 'dark', 'light']) {
		await page.evaluate(
			({ word, mode }) => {
				document.getElementById(word ? 'native-host' : 'editor').dataset.officeTheme = mode;
			},
			{ word, mode },
		);
		colors.push(await toolbar.evaluate((element) => getComputedStyle(element).backgroundColor));
		await expect(bold).toHaveAttribute('aria-pressed', 'true');
		assert(await editor.evaluate((element, node) => element === node, original));
	}
	assert.notEqual(colors[0], colors[1]);
	assert.equal(colors[0], colors[2]);
	console.log(
		`PASS ${product}/${framework}: built native composition, formatting and scoped themes`,
	);
}

try {
	for (const [product, check] of Object.entries({
		docx: word,
		xlsx: excel,
		visio,
		pptx: powerpoint,
	})) {
		for (const framework of frameworks) {
			if (product === 'pptx' && framework === 'solid') continue;
			const route =
				framework === (product === 'visio' ? 'vanilla' : 'react') ? 'demo' : `demo-${framework}`;
			const context = await browser.newContext({
				viewport: { width: 1440, height: 960 },
				colorScheme: 'light',
			});
			const page = await context.newPage();
			const errors = [];
			page.on('pageerror', (error) => errors.push(error.message));
			page.on('response', (response) => {
				if (response.url().startsWith(base) && response.status() >= 400)
					errors.push(`${response.status()} ${response.url()}`);
			});
			page.on('request', (request) => {
				if (request.url().startsWith('https://christophervr.github.io/ooxml/'))
					errors.push(`Preview requested the published site: ${request.url()}`);
			});
			try {
				// No framework query: each production route must select its own compiled default.
				await page.goto(
					`${base}/${product}/${route}/${['pptx', 'visio'].includes(product) ? '?sample=1&locale=en' : ''}`,
				);
				await check(page, framework);
				if (['docx', 'xlsx'].includes(product) && framework !== 'vanilla')
					await nativeBlocks(page, product, framework, route);
				assert.deepEqual(errors, []);
				console.log(`PASS ${product}/${framework}: interaction, themes and document preservation`);
			} catch (error) {
				failures.push(`${product}/${framework}: ${error.message}; ${errors.join('; ')}`);
				console.error(`FAIL ${failures.at(-1)}`);
			} finally {
				await context.close();
			}
		}
	}
} finally {
	await browser.close();
}
assert.equal(failures.length, 0, failures.join('\n'));
console.log('All 23 built framework demos and 10 native compositions passed.');
