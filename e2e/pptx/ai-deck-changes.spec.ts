/* oxlint-disable vitest/prefer-importing-vitest-globals -- Playwright spec, `test`/`expect` come from @playwright/test */
/**
 * What an assistant change does to the deck, through the real panel, in every binding: a theme
 * change re-colours the slides, saves and undoes as one step, a staged change survives Slide
 * Master view edits and slide paste and still applies, and nothing the assistant wrote or staged
 * on one presentation reaches the next one opened.
 *
 * The model is mocked (`support/ai-panel`); the deck is read back through the assistant's own
 * read tools and the bytes a save produces.
 */
import { readFile, writeFile } from 'node:fs/promises';

import { expect, test } from '@playwright/test';
import type { Locator, Page } from '@playwright/test';
import JSZip from 'jszip';
import type { PowerPointViewerAPI } from 'pptx-viewer-shared';

import type { MockToolCall } from './support/ai-panel';
import {
	aiPanel,
	appears,
	installScriptedAiTools,
	latestToolResults,
	openAiPanel,
	seedMockAiProvider,
	sendAiMessage,
} from './support/ai-panel';
import { fixture, loadDeckAt, SAMPLE_DECK } from './support/deck';
import { byBinding } from './support/menu-report';
import { acrossFrameworks } from './support/parity';

declare global {
	interface Window {
		/** Public component handle exposed by each development demo. */
		__pptxViewer: PowerPointViewerAPI;
	}
}

const VIEWPORT = { width: 1600, height: 950 };
test.use({ viewport: VIEWPORT });
test.describe.configure({ timeout: 300_000 });

const TURN_TIMEOUT_MS = 40_000;
const MARK = 'a1111';
const TITLE = 'DECKONETITLE';
const TEXT = 'DECKONETEXT';

const THEME_CHANGE: MockToolCall[] = [
	{
		name: 'update_theme_colors',
		input: Object.fromEntries(
			[
				'accent1',
				'accent2',
				'accent3',
				'accent4',
				'accent5',
				'accent6',
				'dk1',
				'dk2',
				'lt1',
				'lt2',
			].map((key, index) => [key, `#${MARK}${index.toString(16)}`]),
		),
	},
];
const READS: MockToolCall[] = [
	{ name: 'get_metadata', input: {} },
	{ name: 'export_to_json', input: { pretty: false } },
];

interface Assistant {
	panel: Locator;
	elementId: string;
	scripts: Record<string, MockToolCall[]>;
	exchange(message: string): Promise<void>;
	/** What the assistant's read tools report: the theme mark and the title. */
	read(): Promise<{ theme: boolean; title: boolean }>;
	pendingCount(): Promise<number>;
	/** Whether the panel says staged changes were discarded by a load. */
	discardedNotice(): Promise<boolean>;
	applyFirst(): Promise<void>;
	/** Reopen the panel when a view change closed it. */
	ensureOpen(): Promise<void>;
}

async function openAssistant(page: Page, origin: string, deck: string): Promise<Assistant> {
	await seedMockAiProvider(page);
	const scripts: Record<string, MockToolCall[]> = { zzread: READS, zztheme: THEME_CHANGE };
	const log = await installScriptedAiTools(page, scripts);
	await loadDeckAt(page, `${origin}?aiPolicy=stage`, deck);
	const elementId =
		(await page
			.locator('[data-pptx-viewport] [data-element-id]')
			.first()
			.getAttribute('data-element-id')) ?? '';
	await openAiPanel(page);
	const panel = aiPanel(page);
	let turns = 0;
	const exchange = async (message: string): Promise<void> => {
		turns += 1;
		const sent = `${message} ${turns}`;
		await sendAiMessage(page, sent);
		await panel.getByText(`Finished ${sent}`).first().waitFor({ timeout: TURN_TIMEOUT_MS });
		await page.waitForTimeout(500);
	};
	return {
		panel,
		elementId,
		scripts,
		exchange,
		async read() {
			await exchange('zzread');
			const seen = latestToolResults(log).toLowerCase();
			return { theme: seen.includes(MARK), title: seen.includes(TITLE.toLowerCase()) };
		},
		async pendingCount() {
			const label = panel.getByText(/Suggested changes \(\d+\)/u).first();
			if (!(await appears(label, 1_500))) {
				return 0;
			}
			return Number(/\((\d+)\)/u.exec((await label.textContent()) ?? '')?.[1]);
		},
		discardedNotice: () => appears(panel.getByTestId('pptx-ai-proposals-discarded').first(), 1_500),
		async applyFirst() {
			await panel
				.getByRole('button', { name: /^apply$/iu })
				.first()
				.click();
			await page.waitForTimeout(800);
		},
		async ensureOpen() {
			if (!(await appears(panel.locator('textarea').first(), 1_500))) {
				await openAiPanel(page);
			}
		},
	};
}

/** Whether the slides hold a colour from {@link THEME_CHANGE}. */
function slidesRecoloured(page: Page): Promise<boolean> {
	return page.evaluate(
		(mark) => JSON.stringify(window.__pptxViewer.getSlides()).toLowerCase().includes(mark),
		MARK,
	);
}

/** What a save writes: the theme mark in a theme part, the title in the core properties. */
async function saved(page: Page): Promise<{ theme: boolean; title: boolean; text: boolean }> {
	const encoded = await page.evaluate(async () => {
		const bytes = await window.__pptxViewer.getContent();
		let binary = '';
		for (let i = 0; i < bytes.length; i += 0x8000) {
			binary += String.fromCharCode(...bytes.subarray(i, i + 0x8000));
		}
		return btoa(binary);
	});
	const zip = await JSZip.loadAsync(Buffer.from(encoded, 'base64'));
	const parts = async (pattern: RegExp): Promise<string> =>
		(
			await Promise.all(
				Object.keys(zip.files)
					.filter((name) => pattern.test(name))
					.map((name) => zip.file(name)?.async('string') ?? ''),
			)
		).join('');
	return {
		theme: (await parts(/^ppt\/theme\/theme\d*\.xml$/u)).toLowerCase().includes(MARK),
		title: (await parts(/^docProps\/core\.xml$/u)).includes(TITLE),
		text: (await parts(/^ppt\/slides\/slide\d+\.xml$/u)).includes(TEXT),
	};
}

/** Open `path` through File > Open, the way a user replaces the deck. */
async function openThroughFileTab(page: Page, path: string, slides: number): Promise<void> {
	await page
		.getByRole('toolbar', { name: 'Presentation toolbar' })
		.getByRole('tab', { name: 'File', exact: true })
		.click();
	const backstage = page.locator('[role="dialog"][aria-label="File"]');
	await backstage
		.getByRole('button', { name: /^open$/iu })
		.first()
		.click();
	const [chooser] = await Promise.all([
		page.waitForEvent('filechooser', { timeout: 15_000 }),
		backstage
			.getByRole('button', { name: /browse this device/iu })
			.first()
			.click(),
	]);
	await chooser.setFiles(path);
	await page.waitForFunction(
		(count) => window.__pptxViewer?.getSlides?.().length === count,
		slides,
		{
			timeout: 40_000,
		},
	);
	await page.waitForTimeout(1_500);
	await backstage.waitFor({ state: 'hidden', timeout: 5_000 }).catch(() => undefined);
}

/** The Slide Master view sidebar, found through its master tablist. */
function masterSidebar(page: Page): Locator {
	return page
		.getByRole('tablist')
		.filter({ has: page.getByRole('tab', { name: 'Handout', exact: true }) })
		.first()
		.locator('..');
}

/**
 * Insert a layout, then a slide master, then duplicate that master in Slide Master view, and close
 * it. Each command reloads the package into a new handler; none of them loads a presentation.
 */
async function editSlideMasters(page: Page): Promise<void> {
	await page.getByRole('tab', { name: 'View', exact: true }).first().click();
	await page
		.getByRole('toolbar')
		.first()
		.getByRole('button', { name: 'Slide Master', exact: true })
		.click();
	const sidebar = masterSidebar(page);
	await sidebar.waitFor({ timeout: 15_000 });
	const rail = sidebar.locator('button:not([role="tab"]):not([data-testid^="pptx-master-crud-"])');
	for (const command of ['addLayout', 'addMaster', 'duplicateMaster']) {
		const before = await rail.count();
		await sidebar.getByTestId(`pptx-master-crud-${command}`).click();
		await expect(rail).not.toHaveCount(before, { timeout: 15_000 });
	}
	await sidebar.getByRole('button').first().click();
	await expect(masterSidebar(page)).toHaveCount(0);
}

/** Copy the first slide to the system clipboard in Slide Sorter (View tab) and paste it. */
async function pasteSlide(page: Page): Promise<void> {
	await page.context().grantPermissions(['clipboard-read', 'clipboard-write']);
	await page.getByRole('tab', { name: 'View', exact: true }).first().click();
	await page
		.getByRole('toolbar')
		.first()
		.getByRole('button', { name: 'Slide Sorter', exact: true })
		.click();
	const tiles = page.locator('[data-pptx-chrome="sorter-tile"]');
	const before = await tiles.count();
	await tiles.nth(0).click();
	await page.getByRole('button', { name: 'Copy slides to system clipboard', exact: true }).click();
	await page.getByText('Slides copied', { exact: true }).first().waitFor({ timeout: 10_000 });
	await page
		.getByRole('button', { name: 'Paste slides from system clipboard', exact: true })
		.click();
	await expect(tiles).toHaveCount(before + 1, { timeout: 20_000 });
	await page.keyboard.press('Escape');
	await expect(tiles).toHaveCount(0);
}

/** Whether the slides on screen hold the staged text. */
function slidesHoldText(page: Page): Promise<boolean> {
	return page.evaluate(
		(text) => JSON.stringify(window.__pptxViewer.getSlides()).includes(text),
		TEXT,
	);
}

/** `source` without `docProps/core.xml`, so the deck has no title for a stale one to hide behind. */
async function withoutCoreProperties(source: string, out: string): Promise<string> {
	const zip = await JSZip.loadAsync(await readFile(source));
	zip.remove('docProps/core.xml');
	const types = (await zip.file('[Content_Types].xml')?.async('string')) ?? '';
	zip.file('[Content_Types].xml', types.replace(/<Override[^>]*docProps\/core\.xml[^>]*\/>/u, ''));
	const rels = (await zip.file('_rels/.rels')?.async('string')) ?? '';
	zip.file('_rels/.rels', rels.replace(/<Relationship [^>]*core-properties[^>]*\/>/u, ''));
	await writeFile(out, await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' }));
	return out;
}

async function slideCount(path: string): Promise<number> {
	const zip = await JSZip.loadAsync(await readFile(path));
	return Object.keys(zip.files).filter((name) => /^ppt\/slides\/slide\d+\.xml$/u.test(name)).length;
}

function expectAgreement<T>(results: Parameters<typeof byBinding<T>>[0], expected: T): void {
	const problems = byBinding(results).flatMap(({ name, value }) =>
		JSON.stringify(value) === JSON.stringify(expected) ? [] : [`${name}: ${JSON.stringify(value)}`],
	);
	expect(problems.join('\n'), `expected ${JSON.stringify(expected)}`).toBe('');
}

test.describe('cross-binding AI deck changes', () => {
	test(
		'a theme change re-colours the slides, saves and undoes as one step',
		{ tag: '@cross-binding' },
		async ({ browser }, info) => {
			const results = await acrossFrameworks(
				browser,
				info,
				async (page, origin) => {
					const assistant = await openAssistant(page, origin, SAMPLE_DECK);
					const before = await slidesRecoloured(page);
					await assistant.exchange('zztheme');
					await assistant.applyFirst();
					await expect.poll(() => slidesRecoloured(page)).toBe(true);
					const recoloured = await slidesRecoloured(page);
					const read = await assistant.read();
					const undoAll = async (): Promise<number> => {
						let steps = 0;
						while (steps < 5 && (await page.evaluate(() => window.__pptxViewer.canUndo()))) {
							await page.evaluate(() => window.__pptxViewer.undo());
							await page.waitForTimeout(500);
							steps += 1;
						}
						return steps;
					};
					const steps = await undoAll();
					await expect.poll(() => slidesRecoloured(page)).toBe(false);
					const undone = await slidesRecoloured(page);
					const readUndone = await assistant.read();
					await page.evaluate(() => window.__pptxViewer.redo());
					await expect.poll(() => slidesRecoloured(page)).toBe(true);
					const redone = await slidesRecoloured(page);
					const save = await saved(page);
					await undoAll();
					const saveUndone = await saved(page);
					return {
						before,
						recoloured,
						read: read.theme,
						steps,
						undone,
						readUndone: readUndone.theme,
						redone,
						saved: save.theme,
						savedUndone: saveUndone.theme,
					};
				},
				{ concurrency: 'sequential', viewport: VIEWPORT },
			);
			expectAgreement(results, {
				before: false,
				recoloured: true,
				read: true,
				steps: 1,
				undone: false,
				readUndone: false,
				redone: true,
				saved: true,
				savedUndone: false,
			});
		},
	);

	test(
		'nothing written or staged on one deck reaches the next one opened',
		{ tag: '@cross-binding' },
		async ({ browser }, info) => {
			const first = await withoutCoreProperties(SAMPLE_DECK, info.outputPath('first.pptx'));
			const second = await withoutCoreProperties(
				fixture('master-views.pptx'),
				info.outputPath('second.pptx'),
			);
			const secondSlides = await slideCount(second);
			const results = await acrossFrameworks(
				browser,
				info,
				async (page, origin) => {
					const assistant = await openAssistant(page, origin, first);
					assistant.scripts.zztitle = [{ name: 'update_metadata', input: { title: TITLE } }];
					assistant.scripts.zztext = [
						{
							name: 'update_element',
							input: { slideIndex: 0, elementId: assistant.elementId, text: TEXT },
						},
					];
					await assistant.exchange('zztitle');
					await assistant.applyFirst();
					await assistant.exchange('zztext');
					const staged = await assistant.pendingCount();
					// A render passes with no assistant read between the write and the next deck.
					await page.waitForTimeout(1_200);
					await openThroughFileTab(page, second, secondSlides);
					await assistant.ensureOpen();
					const left = await assistant.pendingCount();
					const notice = await assistant.discardedNotice();
					const read = await assistant.read();
					const save = await saved(page);
					return { staged, left, notice, title: read.title, saved: save };
				},
				{ concurrency: 'sequential', viewport: VIEWPORT },
			);
			expectAgreement(results, {
				staged: 1,
				left: 0,
				notice: true,
				title: false,
				saved: { theme: false, title: false, text: false },
			});
		},
	);

	test(
		'a staged change survives Slide Master edits and slide paste, then applies',
		{ tag: '@cross-binding' },
		async ({ browser }, info) => {
			const results = await acrossFrameworks(
				browser,
				info,
				async (page, origin) => {
					const assistant = await openAssistant(page, origin, fixture('master-views.pptx'));
					assistant.scripts.zztext = [
						{
							name: 'update_element',
							input: { slideIndex: 0, elementId: assistant.elementId, text: TEXT },
						},
					];
					await assistant.exchange('zztext');
					const staged = await assistant.pendingCount();
					await editSlideMasters(page);
					await assistant.ensureOpen();
					const afterMasters = await assistant.pendingCount();
					await pasteSlide(page);
					await assistant.ensureOpen();
					const afterPaste = await assistant.pendingCount();
					const notice = await assistant.discardedNotice();
					await assistant.applyFirst();
					const conflict = await assistant.panel
						.getByText(/changed after this was proposed/iu)
						.count();
					const left = await assistant.pendingCount();
					const applied = await slidesHoldText(page);
					const save = await saved(page);
					return {
						staged,
						afterMasters,
						afterPaste,
						notice,
						conflict,
						left,
						applied,
						saved: save.text,
					};
				},
				{ concurrency: 'sequential', viewport: VIEWPORT },
			);
			expectAgreement(results, {
				staged: 1,
				afterMasters: 1,
				afterPaste: 1,
				notice: false,
				conflict: 0,
				left: 0,
				applied: true,
				saved: true,
			});
		},
	);
});
