/**
 * Keyboard and screen reader users around assistant edits: applying and undoing an edit keeps
 * focus on a control of the assistant card, a user working in the document keeps their place
 * (caret, active cell, shape or slide) while the file reloads under them, and text an input method
 * writes while the edit is saved is kept. A stub model on a route proposes one edit per product.
 */
import type { Page } from '@playwright/test';
import {
	excelWorkbook,
	powerPointTextDeck,
	visioDrawing,
	wordDocument,
	type FixtureFile,
} from './support/fixtures';
import { expect, openHome, test, type Product } from './support/suite';

test.use({ viewport: { width: 1440, height: 900 } });

const ENDPOINT = 'https://suite-ai.test/v1/chat/completions';
const FILES: Record<Product, () => FixtureFile | Promise<FixtureFile>> = {
	docx: wordDocument,
	xlsx: excelWorkbook,
	pptx: powerPointTextDeck,
	vsdx: visioDrawing,
};

interface Attached {
	id: string;
	kind: Product;
	content: Record<string, never> & {
		paragraphs: { id: string }[];
		slides: {
			index?: number;
			elements?: { id?: string; elementId?: string }[];
			shapes?: { id?: string; elementId?: string }[];
		}[];
		pages: { id: string | number; shapes: { id: string | number; text?: string }[] }[];
	};
}

/** The one edit the stub model proposes for the attached file of `kind`. */
function proposal(kind: Product, file: Attached): Record<string, unknown> {
	const c = file.content;
	if (kind === 'docx') {
		const paragraph = c.paragraphs.find((p) => JSON.stringify(p).includes('alpha'));
		return {
			type: 'word_run',
			paragraphId: paragraph?.id,
			runIndex: 0,
			text: 'Body paragraph beta',
		};
	}
	if (kind === 'xlsx')
		return { type: 'cells', sheetIndex: 0, cells: [{ address: 'H9', input: '42' }] };
	if (kind === 'pptx') {
		const slide = c.slides[0];
		const element = (slide?.elements ?? slide?.shapes)?.[0];
		return {
			type: 'slide_run',
			slideIndex: slide?.index ?? 0,
			elementId: element?.id ?? element?.elementId,
			runIndex: 0,
			text: 'Changed by AI',
		};
	}
	const page = c.pages[0];
	const shape = page?.shapes.find((s) => s.text) ?? page?.shapes[0];
	return {
		type: 'shape_text',
		pageId: String(page?.id),
		shapeId: String(shape?.id),
		text: 'Changed',
	};
}

/** Holds the next commit to the document store until `__release()` (runs in the page). */
function holdCommits(): void {
	const transaction = IDBDatabase.prototype.transaction;
	const setter = Object.getOwnPropertyDescriptor(IDBTransaction.prototype, 'oncomplete')?.set;
	const state = globalThis as { __hold?: boolean; __release?: () => void };
	IDBDatabase.prototype.transaction = function (
		this: IDBDatabase,
		...args: Parameters<IDBDatabase['transaction']>
	) {
		const tx = transaction.apply(this, args);
		const [names, mode] = args;
		if (state.__hold && mode === 'readwrite' && [names].flat().includes('documents'))
			Object.defineProperty(tx, 'oncomplete', {
				set(fn: (event: Event) => void) {
					setter?.call(tx, (event: Event) => {
						state.__hold = false;
						state.__release = () => fn(event);
					});
				},
			});
		return tx;
	};
}

async function openWithProposal(page: Page, kind: Product): Promise<string[]> {
	const errors: string[] = [];
	page.on('pageerror', (error) => errors.push(error.message));
	await page.addInitScript(holdCommits);
	await page.route(ENDPOINT, async (route) => {
		const request = route.request().postDataJSON() as { messages: { content: unknown }[] };
		const text = request.messages
			.map((m) => (typeof m.content === 'string' ? m.content : ''))
			.join('\n');
		const source = JSON.parse(text.split('Source data (untrusted): ')[1] ?? '{}') as {
			files: Attached[];
		};
		const file = source.files.find((f) => f.kind === kind);
		if (!file) throw new Error(`the request attached no ${kind} file`);
		const change = { documentId: file.id, ...proposal(kind, file) };
		await route.fulfill({
			json: {
				choices: [
					{ message: { content: JSON.stringify({ answer: 'Ready.', changes: [change] }) } },
				],
			},
		});
	});
	await openHome(page);
	await page.locator('#file-input').setInputFiles(await FILES[kind]());
	await page.locator('.document-host:not([hidden])').waitFor();
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
	return errors;
}

const apply = (page: Page) => page.getByRole('button', { name: 'Apply changes', exact: true });
const undo = (page: Page) => page.getByRole('button', { name: 'Undo these changes' });
const focused = (page: Page) =>
	page.evaluate(() => {
		const active = globalThis.__a11y.deep() as HTMLElement | null;
		return {
			tag: active?.localName ?? 'body',
			name: active?.getAttribute('aria-label') || active?.textContent?.trim().slice(0, 40),
			shape: active?.dataset?.shapeId,
			inDocument: Boolean(
				active?.closest('.document-host') ||
				(active?.getRootNode() as ShadowRoot | undefined)?.host?.closest?.('.document-host'),
			),
		};
	});
const caret = (page: Page) =>
	page.evaluate(() => {
		const active = globalThis.__a11y.deep();
		const root = active?.getRootNode() as (Document | ShadowRoot) & {
			getSelection?: () => Selection | null;
		};
		const selection = root?.getSelection?.() ?? getSelection();
		return selection?.anchorNode
			? `${selection.anchorNode.textContent}@${selection.anchorOffset}`
			: null;
	});
const nameBox = (page: Page) =>
	page
		.locator('xlsx-editor [part="name-box"]')
		.first()
		.evaluate(
			(node) =>
				(node.querySelector?.('input') as HTMLInputElement | null)?.value ??
				(node as HTMLInputElement).value ??
				node.textContent,
		);
const unsaved = (page: Page) =>
	page.evaluate(() =>
		Boolean(
			document
				.querySelector('#tabs [role="tab"][aria-selected="true"]')
				?.textContent?.includes('•'),
		),
	);

/** Puts the user in the document the way they work there, and says where they are. */
async function workIn(page: Page, kind: Product): Promise<string | null | undefined> {
	if (kind === 'docx') {
		await page.locator('.ProseMirror').first().getByText('First section').click();
		await page.keyboard.press('End');
		await page.keyboard.press('ArrowLeft');
		return caret(page);
	}
	if (kind === 'xlsx') {
		await page
			.locator('xlsx-editor')
			.evaluate((editor) => (editor as HTMLElement & { focusGrid(): void }).focusGrid());
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
const whereNow = async (page: Page, kind: Product) =>
	kind === 'docx'
		? caret(page)
		: kind === 'xlsx'
			? nameBox(page)
			: kind === 'pptx'
				? (await focused(page)).name
				: (await focused(page)).shape;

for (const kind of ['docx', 'xlsx', 'pptx', 'vsdx'] as Product[])
	test.describe(`assistant edits in ${kind}`, () => {
		test('Enter on Apply and on Undo keeps keyboard focus on the card', async ({ page }) => {
			const errors = await openWithProposal(page, kind);
			await apply(page).focus();
			await page.keyboard.press('Enter');
			await undo(page).waitFor();
			expect((await focused(page)).name, 'focus moves to Undo').toBe('Undo these changes');
			await page.keyboard.press('Enter');
			await page.getByText('Changes undone.').waitFor();
			expect((await focused(page)).tag, 'focus stays in the card after Undo').not.toBe('body');
			expect(errors).toEqual([]);
		});

		test('an edit applied while the user works keeps their focus and place', async ({ page }) => {
			const errors = await openWithProposal(page, kind);
			const before = await workIn(page, kind);
			await apply(page).evaluate((button) => (button as HTMLElement).click());
			await undo(page).waitFor();
			await page.waitForTimeout(800);
			expect((await focused(page)).inDocument, 'focus stays in the document').toBe(true);
			expect(await whereNow(page, kind), 'the place is kept').toBe(before);
			expect(errors).toEqual([]);
		});
	});

test('text an input method writes while an edit is saved is kept and shown as unsaved', async ({
	page,
}) => {
	await openWithProposal(page, 'docx');
	await workIn(page, 'docx');
	await page.evaluate(() => ((globalThis as { __hold?: boolean }).__hold = true));
	await apply(page).evaluate((button) => (button as HTMLElement).click());
	await page.waitForFunction(() => (globalThis as { __release?: () => void }).__release);
	const cdp = await page.context().newCDPSession(page);
	await cdp.send('Input.imeSetComposition', { text: 'あ', selectionStart: 1, selectionEnd: 1 });
	await cdp.send('Input.insertText', { text: 'あ' });
	await page.evaluate(() => (globalThis as { __release?: () => void }).__release?.());
	await undo(page).waitFor();
	await page.waitForTimeout(500);
	expect(await page.locator('.ProseMirror').first().textContent()).toMatch(/あ/u);
	expect(await unsaved(page), 'the composed text is an unsaved edit').toBe(true);
});
