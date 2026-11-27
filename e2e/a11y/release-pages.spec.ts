/**
 * The release posts (`releases/<version>/`, built into suite-dist by build:suite). Both keep a
 * stored `ooxml-suite-theme` choice, toggle it from a named switch and keep their text at 4.5:1
 * in both themes. 2.0 follows the device until the reader chooses. 2.5 wears its own edition: it
 * opens dark, on black, and moves as it is read, so these also check that the motion never
 * costs the reader anything: the text is always in the accessibility tree, a link into the page
 * lands on something visible, reduced motion gets a still page, and axe finds nothing critical
 * or serious in light, in dark and on a phone.
 */
import { expect, test, type BrowserContext, type Page } from '@playwright/test';
import { axeViolations } from './support/axe';

const POSTS = ['2.0', '2.5'] as const;
const KEY = 'ooxml-suite-theme';
const PHONE = { width: 390, height: 844 };

/** Text that must stay readable, per post: the shared article and each edition's own parts. */
const TEXT: Record<(typeof POSTS)[number], readonly string[]> = {
	'2.0': ['.post-hero h1', '.post-meta', '.post-article p', '.post-article a[href]', '.post-toc a'],
	'2.5': [
		'.post-brand',
		'.post-bar-links a',
		'.post-icon',
		'.post-hero h1',
		'.post-meta',
		'.release-intro',
		'.index-name',
		'.index-links a',
		'.index-rest .index-name',
		'.post-article p',
		'.post-article a[href]',
		'.chapter-head h2',
		'.chapter-links a',
		'.post-article .chapter-head + p',
		'.feature > h3',
		'.post-figure figcaption',
		'.recording-bar button',
		'.recording-bar a',
		'.not-yet',
		'.not-yet strong',
		'.learn-more',
		'.code-file',
		'.changes-name',
		'.changes-count',
		'.post-footer',
	],
};

/** Contrast of text against the nearest painted background, as the browser resolved them. */
async function contrastOf(page: Page, selector: string): Promise<number> {
	return page
		.locator(selector)
		.first()
		.evaluate((element) => {
			const rgb = (value: string) =>
				(/rgba?\(([^)]+)\)/u.exec(value)?.[1] ?? '0, 0, 0')
					.split(',')
					.slice(0, 3)
					.map((part) => Number(part.trim()) / 255);
			const luminance = (value: string) => {
				const [r = 0, g = 0, b = 0] = rgb(value).map((c) =>
					c <= 0.03928 ? c / 12.92 : ((c + 0.055) / 1.055) ** 2.4,
				);
				return 0.2126 * r + 0.7152 * g + 0.0722 * b;
			};
			let paint: Element | null = element;
			let background = 'rgb(255, 255, 255)';
			while (paint) {
				const fill = getComputedStyle(paint).backgroundColor;
				if (fill !== 'rgba(0, 0, 0, 0)' && fill !== 'transparent') {
					background = fill;
					break;
				}
				paint = paint.parentElement;
			}
			const [a, b] = [luminance(getComputedStyle(element).color), luminance(background)];
			return (Math.max(a, b) + 0.05) / (Math.min(a, b) + 0.05);
		});
}

/** The theme a post opens in when nothing is stored: 2.0 follows the device, 2.5 opens dark. */
const opensIn = (post: (typeof POSTS)[number], device: 'light' | 'dark') =>
	post === '2.5' ? 'dark' : device;
const other = (theme: string) => (theme === 'dark' ? 'light' : 'dark');
/** A stored choice, set before the page loads, as the suite or an earlier visit would leave it. */
async function store(context: BrowserContext, mode: 'light' | 'dark' | 'system'): Promise<void> {
	await context.addInitScript(
		({ key, mode }) => {
			try {
				localStorage.setItem(key, JSON.stringify({ mode, accent: 'blue' }));
			} catch {
				// A page without storage (about:blank) has nothing to remember.
			}
		},
		{ key: KEY, mode },
	);
}
/** How many parts of the 2.5 page are waiting, unseen, for their place in the window. */
const waiting = (page: Page) =>
	page.evaluate(
		() =>
			[...document.querySelectorAll('[data-rv]')].filter((part) => !part.classList.contains('in'))
				.length,
	);

for (const post of POSTS) {
	for (const device of ['light', 'dark'] as const) {
		test(`${post} on a ${device} device opens in its own default`, async ({ browser }) => {
			const context = await browser.newContext({ colorScheme: device });
			const page = await context.newPage();
			await page.goto(`/releases/${post}/`);
			const theme = opensIn(post, device);
			await expect(page.locator('html')).toHaveAttribute('data-theme', theme);
			await expect(page.getByRole('button', { name: 'Dark mode' })).toHaveAttribute(
				'aria-pressed',
				String(theme === 'dark'),
			);
			// Nothing is stored until the reader chooses.
			expect(await page.evaluate((key) => localStorage.getItem(key), KEY)).toBeNull();
			await context.close();
		});

		test(`${post} keeps a stored choice on a ${device} device`, async ({ browser }) => {
			for (const mode of ['light', 'dark', 'system'] as const) {
				const context = await browser.newContext({ colorScheme: device });
				await store(context, mode);
				const page = await context.newPage();
				await page.goto(`/releases/${post}/`);
				await expect(page.locator('html'), `stored ${mode}`).toHaveAttribute(
					'data-theme',
					mode === 'system' ? device : mode,
				);
				await context.close();
			}
		});
	}

	for (const theme of ['light', 'dark'] as const)
		test(`${post} keeps text readable in ${theme}`, async ({ browser }) => {
			const context = await browser.newContext({ reducedMotion: 'reduce' });
			await store(context, theme);
			const page = await context.newPage();
			await page.goto(`/releases/${post}/`);
			await expect(page.locator('html')).toHaveAttribute('data-theme', theme);
			for (const selector of TEXT[post])
				expect(await contrastOf(page, selector), `${selector} in ${theme}`).toBeGreaterThanOrEqual(
					4.5,
				);
			await context.close();
		});

	test(`${post}: the switch toggles the theme and remembers the choice`, async ({ browser }) => {
		const context = await browser.newContext({ colorScheme: 'light' });
		const page = await context.newPage();
		await page.goto(`/releases/${post}/`);
		const first = opensIn(post, 'light');
		const toggle = page.getByRole('button', { name: 'Dark mode' });
		await toggle.focus();
		await page.keyboard.press('Enter');
		await expect(page.locator('html')).toHaveAttribute('data-theme', other(first));
		await expect(toggle).toHaveAttribute('aria-pressed', String(other(first) === 'dark'));
		expect(
			await page.evaluate((key) => JSON.parse(localStorage.getItem(key) ?? '{}').mode, KEY),
		).toBe(other(first));
		await page.reload();
		await expect(page.locator('html')).toHaveAttribute('data-theme', other(first));
		await page.getByRole('button', { name: 'Dark mode' }).click();
		await expect(page.locator('html')).toHaveAttribute('data-theme', first);
		await context.close();
	});
}

for (const scheme of ['light', 'dark'] as const)
	for (const [layout, viewport] of [
		['desktop', { width: 1440, height: 900 }],
		['phone', PHONE],
	] as const)
		test(`2.5 has no critical or serious axe violations, ${scheme} on a ${layout}`, async ({
			browser,
		}) => {
			// The still page: axe reads what is drawn, and a part mid-entrance is not yet drawn.
			const context = await browser.newContext({ viewport, reducedMotion: 'reduce' });
			await store(context, scheme);
			const page = await context.newPage();
			for (const path of ['', 'migration.html']) {
				await page.goto(`/releases/2.5/${path}`);
				await expect(page.locator('h1')).toBeVisible();
				const found = (await axeViolations(page)).map(
					(v) => `${v.rule} (${v.impact}): ${v.nodes} nodes, e.g. ${v.targets.join(' | ')}`,
				);
				expect(found, `axe on releases/2.5/${path}`).toEqual([]);
				// Nothing is wider than the window: the page never scrolls sideways.
				expect(
					await page.evaluate(() => document.documentElement.scrollWidth - innerWidth),
					`sideways scroll on releases/2.5/${path}`,
				).toBeLessThanOrEqual(0);
			}
			await context.close();
		});

test('the two posts are different editions', async ({ page }) => {
	await page.goto('/releases/2.0/');
	await expect(page.locator('body')).not.toHaveAttribute('data-edition');
	await expect(page.locator('.release-index')).toBeVisible();
	await expect(page.locator('.edition-index')).toHaveCount(0);
	await expect(page.locator('body')).not.toHaveClass(/motion/u);
	await page.goto('/releases/2.5/');
	await expect(page.locator('body')).toHaveAttribute('data-edition', 'letter');
	await expect(page.locator('.release-index, [data-showcase]')).toHaveCount(0);
	// 2.5 is a narrow column on black, set left; 2.0 is centred on the suite's canvas.
	expect(
		await page.locator('body').evaluate((body) => getComputedStyle(body).backgroundColor),
	).toBe('rgb(0, 0, 0)');
	expect(
		await page.locator('.post-hero h1').evaluate((h1) => getComputedStyle(h1).textAlign),
	).not.toBe('center');
	expect(
		await page.locator('.post-article').evaluate((article) => article.clientWidth),
	).toBeLessThanOrEqual(720);
});

test('the 2.5 post is read from its contents and links its migration guide', async ({ page }) => {
	await page.goto('/releases/2.5/');
	await expect(page.locator('.post-hero h1')).toHaveText('OOXML 2.5');
	await expect(page.locator('.post-lede a[href="migration.html"]')).toBeVisible();
	const index = page.getByRole('navigation', { name: 'In OOXML 2.5' });
	await expect(index.locator('.index-row[data-product]')).toHaveCount(5);
	// A line of the contents goes to its entry, which is there to read when the page arrives.
	await index.getByRole('link', { name: 'Equation' }).click();
	const heading = page.getByRole('heading', { level: 3, name: 'Equation' });
	await expect(heading).toBeInViewport();
	await expect(heading).toHaveClass(/\bin\b/u);
	await expect(
		page
			.locator('.feature', { has: page.locator('#equation') })
			.locator('p')
			.first(),
	).toHaveCSS('opacity', '1');
	await page.goto('/releases/2.5/migration.html');
	await expect(page.locator('body')).toHaveAttribute('data-edition', 'letter');
	await expect(page.locator('html')).toHaveAttribute('data-theme', 'dark');
	await expect(
		page.getByRole('heading', { name: 'Word view modes have six values' }),
	).toBeAttached();
	await expect(
		page.getByRole('heading', { name: 'Ink brush widths are written in centimetres' }),
	).toBeAttached();
});

test('the 2.5 page moves as it is read and never hides its text from anyone', async ({
	browser,
}) => {
	const context = await browser.newContext({ reducedMotion: 'no-preference' });
	const page = await context.newPage();
	await page.goto('/releases/2.5/');
	await expect(page.locator('body')).toHaveClass(/motion/u);
	// Far down the page a chapter waits, unseen, but it is in the accessibility tree as itself.
	expect(await waiting(page)).toBeGreaterThan(100);
	const security = page.getByRole('heading', { level: 2, name: 'Security' });
	await expect(security).toHaveCount(1);
	await expect(security).not.toHaveClass(/\bin\b/u);
	expect(await page.locator('#security').evaluate((h2) => h2.textContent)).toBe('Security');
	// Only opacity and transform hide a waiting part: it is laid out and takes its room.
	const paragraph = page.locator('#security').locator('xpath=../following-sibling::ul[1]/li[1]');
	expect(await paragraph.evaluate((li) => getComputedStyle(li).display)).not.toBe('none');
	expect(await paragraph.evaluate((li) => getComputedStyle(li).visibility)).toBe('visible');
	// Reaching it brings it in, and focus never lands on something unseen.
	await security.scrollIntoViewIfNeeded();
	await expect(security).toHaveClass(/\bin\b/u);
	await expect(paragraph).toHaveClass(/\bin\b/u);
	await page.getByRole('link', { name: 'security regression corpus' }).focus();
	await expect(
		page.locator('p', { has: page.getByRole('link', { name: 'security regression corpus' }) }),
	).toHaveClass(/\bin\b/u);
	await context.close();
});

test('a link into the 2.5 page lands on content that is already there', async ({ browser }) => {
	const context = await browser.newContext({ reducedMotion: 'no-preference' });
	const page = await context.newPage();
	await page.goto('/releases/2.5/#excel-draw');
	const heading = page.locator('#excel-draw');
	await expect(heading).toBeInViewport();
	await expect(heading).toHaveClass(/\bin\b/u);
	const entry = page.locator('.feature', { has: heading });
	await expect(entry.locator('p').first()).toHaveCSS('opacity', '1');
	// What the reader skipped over is shown when they scroll back up to it.
	await page.locator('#excel').scrollIntoViewIfNeeded();
	await expect(page.locator('#excel')).toHaveClass(/\bin\b/u);
	await context.close();
});

test('the 2.5 recordings load lazily, carry a caption and can be paused', async ({ page }) => {
	await page.goto('/releases/2.5/');
	const recordings = page.locator('.post-figure img[data-animation$=".gif"]');
	await expect(recordings).toHaveCount(13);
	for (const img of await recordings.all()) {
		expect((await img.getAttribute('alt'))?.length).toBeGreaterThanOrEqual(60);
		expect(await img.getAttribute('loading')).toBe('lazy');
	}
	await expect(page.locator('.post-figure figcaption')).toHaveCount(13);
	// In view, a recording plays; its control pauses it on the poster and plays it again.
	const figure = page.locator('.feature', { has: page.locator('#mailings') }).locator('figure');
	const img = figure.locator('img');
	await img.scrollIntoViewIfNeeded();
	await expect(img).toHaveAttribute('src', /word-mail-merge\.gif$/u);
	await figure.getByRole('button', { name: 'Pause recording' }).click();
	await expect(img).toHaveAttribute('src', /word-mail-merge\.poster\.webp$/u);
	await figure.getByRole('button', { name: 'Play recording' }).click();
	await expect(img).toHaveAttribute('src', /word-mail-merge\.gif$/u);
	await expect(figure.getByRole('link', { name: 'Full size' })).toHaveAttribute(
		'href',
		/word-mail-merge\.gif$/u,
	);
});

test('under reduced motion the 2.5 page is still and its recordings wait to be played', async ({
	browser,
}) => {
	const context = await browser.newContext({ reducedMotion: 'reduce' });
	const page = await context.newPage();
	await page.goto('/releases/2.5/');
	await expect(page.locator('body')).not.toHaveClass(/motion/u);
	expect(await page.locator('[data-rv]').count()).toBe(0);
	for (const selector of ['.post-hero h1', '.release-intro', '.index-row', '.post-lede > p'])
		expect(
			await page
				.locator(selector)
				.first()
				.evaluate((part) => getComputedStyle(part).animationName),
			selector,
		).toBe('none');
	await expect(page.locator('#security')).toHaveCSS('opacity', '1');
	const figure = page.locator('.feature', { has: page.locator('#views') }).locator('figure');
	const img = figure.locator('img');
	await img.scrollIntoViewIfNeeded();
	await expect(figure.getByRole('button', { name: 'Play recording' })).toBeVisible();
	await expect(img).toHaveAttribute('src', /word-views\.poster\.webp$/u);
	await figure.getByRole('button', { name: 'Play recording' }).click();
	await expect(img).toHaveAttribute('src', /word-views\.gif$/u);
	await expect(figure.getByRole('button', { name: 'Pause recording' })).toBeVisible();
	await context.close();
});

test('on a phone the 2.5 contents stack and nothing scrolls sideways', async ({ browser }) => {
	const context = await browser.newContext({ viewport: PHONE, hasTouch: true, isMobile: true });
	const page = await context.newPage();
	await page.goto('/releases/2.5/');
	const row = page.locator('.index-row[data-product="word"]');
	const [name, links] = await Promise.all(
		[row.locator('.index-name'), row.locator('.index-links')].map((part) => part.boundingBox()),
	);
	expect(links!.y).toBeGreaterThanOrEqual(name!.y + name!.height - 1);
	await page.locator('#tables-and-arrange').scrollIntoViewIfNeeded();
	expect(
		await page.evaluate(() => document.documentElement.scrollWidth - innerWidth),
	).toBeLessThanOrEqual(0);
	await context.close();
});
