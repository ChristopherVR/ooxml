/* oxlint-disable vitest/prefer-importing-vitest-globals -- Playwright spec, `test`/`expect` come from @playwright/test */
/**
 * Each assistant write policy, through the real panel, in every binding.
 *
 * The demos read `?aiPolicy=` (`demos/pptx/shared/demo-ai-path.ts`), so these tests use the
 * production path from composer to deck with only the model mocked. What the deck holds is read
 * back the way a model sees it: through the assistant's own read tools.
 */
import { expect, test } from '@playwright/test';
import type { Locator, Page } from '@playwright/test';

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
import { loadDeckAt, SAMPLE_DECK } from './support/deck';
import { byBinding } from './support/menu-report';
import { acrossFrameworks } from './support/parity';

const VIEWPORT = { width: 1600, height: 950 };
test.use({ viewport: VIEWPORT });
test.describe.configure({ timeout: 280_000 });

const TURN_TIMEOUT_MS = 40_000;
const SAMPLE_DECK_SLIDES = 7;

/** One theme, metadata and text write per field, twice over, plus a forced-approval delete. */
function proposedWrites(elementId: string): MockToolCall[] {
	return [
		{ name: 'update_theme_colors', input: { accent1: '#a11111' } },
		{ name: 'update_theme_colors', input: { accent2: '#a22222' } },
		{ name: 'update_metadata', input: { title: 'POLICYTITLE' } },
		{ name: 'update_metadata', input: { creator: 'POLICYAUTHOR' } },
		{ name: 'update_element', input: { slideIndex: 0, elementId, text: 'POLICYTEXT' } },
		{ name: 'delete_slides', input: { slideIndexes: [1] } },
	];
}
const WRITES_BEFORE_DELETE = 5;

const READS: MockToolCall[] = [
	{ name: 'get_metadata', input: {} },
	{ name: 'export_to_json', input: { pretty: false } },
	{ name: 'get_deck_overview', input: {} },
];

interface DeckView {
	accent1: boolean;
	accent2: boolean;
	title: boolean;
	author: boolean;
	text: boolean;
	slides: number;
}

const NOTHING_APPLIED = {
	accent1: false,
	accent2: false,
	title: false,
	author: false,
	text: false,
};
const EVERYTHING_APPLIED = { accent1: true, accent2: true, title: true, author: true, text: true };

interface Assistant {
	panel: Locator;
	/** Send a scripted message and wait for its closing reply. */
	exchange(message: string): Promise<void>;
	readDeck(): Promise<DeckView>;
	pendingCount(): Promise<number>;
}

async function openAssistant(page: Page, origin: string, policy: string): Promise<Assistant> {
	await seedMockAiProvider(page);
	const scripts: Record<string, MockToolCall[]> = { zzread: READS };
	const log = await installScriptedAiTools(page, scripts);
	await loadDeckAt(page, `${origin}?aiPolicy=${policy}`, SAMPLE_DECK);
	const elementId = await page
		.locator('[data-pptx-viewport] [data-element-id]')
		.first()
		.getAttribute('data-element-id');
	scripts.zzpropose = proposedWrites(elementId ?? '');
	await openAiPanel(page);
	const panel = aiPanel(page);
	let reads = 0;

	const exchange = async (message: string): Promise<void> => {
		await sendAiMessage(page, message);
		await panel.getByText(`Finished ${message}`).first().waitFor({ timeout: TURN_TIMEOUT_MS });
	};
	return {
		panel,
		exchange,
		async readDeck() {
			reads += 1;
			await exchange(`zzread ${reads}`);
			const seen = latestToolResults(log).toLowerCase();
			return {
				accent1: seen.includes('a11111'),
				accent2: seen.includes('a22222'),
				title: seen.includes('policytitle'),
				author: seen.includes('policyauthor'),
				text: seen.includes('policytext'),
				slides: Number(/slidecount\\*":\s*(\d+)/u.exec(seen)?.[1]),
			};
		},
		async pendingCount() {
			const label = panel.getByText(/Suggested changes \(\d+\)/u).first();
			if (!(await appears(label, 1_500))) {
				return 0;
			}
			return Number(/\((\d+)\)/u.exec((await label.textContent()) ?? '')?.[1]);
		},
	};
}

function applyButtons(panel: Locator): Locator {
	return panel.getByRole('button', { name: /^apply$/iu });
}

function applyAllButton(panel: Locator): Locator {
	return panel.getByRole('button', { name: /^apply all$/iu });
}

function approvalNotes(panel: Locator): Locator {
	return panel.getByText('Needs your approval');
}

/** A short pause for a click's effects to reach the deck and the panel. */
async function settle(page: Page): Promise<void> {
	await page.waitForTimeout(600);
}

function expectAgreement<T>(results: Parameters<typeof byBinding<T>>[0], expected: T): void {
	const problems = byBinding(results).flatMap(({ name, value }) =>
		JSON.stringify(value) === JSON.stringify(expected) ? [] : [`${name}: ${JSON.stringify(value)}`],
	);
	expect(problems.join('\n'), `expected ${JSON.stringify(expected)}`).toBe('');
}

test.describe('cross-binding AI write policies', () => {
	test(
		'stage: Apply all keeps every change',
		{ tag: '@cross-binding' },
		async ({ browser }, info) => {
			const results = await acrossFrameworks(
				browser,
				info,
				async (page, origin) => {
					const assistant = await openAssistant(page, origin, 'stage');
					await assistant.exchange('zzpropose');
					await settle(page);
					const staged = await assistant.pendingCount();
					const before = await assistant.readDeck();
					await applyAllButton(assistant.panel).first().click();
					await settle(page);
					const left = await assistant.pendingCount();
					const after = await assistant.readDeck();
					await applyButtons(assistant.panel).first().click();
					await settle(page);
					const final = await assistant.readDeck();
					return { staged, before, left, after, finalSlides: final.slides };
				},
				{ concurrency: 'sequential', viewport: VIEWPORT },
			);
			expectAgreement(results, {
				staged: WRITES_BEFORE_DELETE + 1,
				before: { ...NOTHING_APPLIED, slides: SAMPLE_DECK_SLIDES },
				// The delete needs its own approval, so Apply all leaves it.
				left: 1,
				after: { ...EVERYTHING_APPLIED, slides: SAMPLE_DECK_SLIDES },
				finalSlides: SAMPLE_DECK_SLIDES - 1,
			});
		},
	);

	test('auto: changes apply at once', { tag: '@cross-binding' }, async ({ browser }, info) => {
		const results = await acrossFrameworks(
			browser,
			info,
			async (page, origin) => {
				const assistant = await openAssistant(page, origin, 'auto');
				await assistant.exchange('zzpropose');
				await settle(page);
				const pending = await assistant.pendingCount();
				const notes = await approvalNotes(assistant.panel).count();
				const deck = await assistant.readDeck();
				await applyButtons(assistant.panel).first().click();
				await settle(page);
				const final = await assistant.readDeck();
				return { pending, notes, deck, finalSlides: final.slides };
			},
			{ concurrency: 'sequential', viewport: VIEWPORT },
		);
		expectAgreement(results, {
			// Only the delete waits: it asks for approval under every policy.
			pending: 1,
			notes: 1,
			deck: { ...EVERYTHING_APPLIED, slides: SAMPLE_DECK_SLIDES },
			finalSlides: SAMPLE_DECK_SLIDES - 1,
		});
	});

	test(
		'approve: each change is accepted by itself',
		{ tag: '@cross-binding' },
		async ({ browser }, info) => {
			const results = await acrossFrameworks(
				browser,
				info,
				async (page, origin) => {
					const assistant = await openAssistant(page, origin, 'approve');
					await assistant.exchange('zzpropose');
					await settle(page);
					const pending = await assistant.pendingCount();
					const notes = await approvalNotes(assistant.panel).count();
					const applyAll = await applyAllButton(assistant.panel).count();
					const before = await assistant.readDeck();
					for (let accepted = 0; accepted < pending; accepted += 1) {
						await applyButtons(assistant.panel).first().click();
						await settle(page);
					}
					const left = await assistant.pendingCount();
					const after = await assistant.readDeck();
					return { pending, notes, applyAll, before, left, after };
				},
				{ concurrency: 'sequential', viewport: VIEWPORT },
			);
			expectAgreement(results, {
				pending: WRITES_BEFORE_DELETE + 1,
				notes: WRITES_BEFORE_DELETE + 1,
				applyAll: 0,
				before: { ...NOTHING_APPLIED, slides: SAMPLE_DECK_SLIDES },
				left: 0,
				after: { ...EVERYTHING_APPLIED, slides: SAMPLE_DECK_SLIDES - 1 },
			});
		},
	);
});
