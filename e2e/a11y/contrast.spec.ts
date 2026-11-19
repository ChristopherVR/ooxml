/**
 * Token contrast for every accent in light and dark: the pairs the suite and its editors paint
 * text and indicators with, read as the browser resolves them. Text needs 4.5:1 (WCAG 1.4.3);
 * focus rings, selection indicators and control borders need 3:1 (1.4.11).
 */
import { gate } from './support/known';
import { ACCENTS, expect, openHome, test, type Accent, type Mode } from './support/suite';

/** A pair: the colour painted, what it is painted on, and the ratio it needs. */
const PAIRS = {
	'text on background': ['--office-foreground', '--office-background', 4.5],
	'text on surface': ['--office-foreground', '--office-surface', 4.5],
	'secondary text on background': ['--office-muted-foreground', '--office-background', 4.5],
	'secondary text on surface': ['--office-muted-foreground', '--office-surface', 4.5],
	'text on accent': ['--office-accent-foreground', '--office-accent', 4.5],
	'accent text on background': ['--accent-ink', '--office-background', 4.5],
	'Teams brand button text': ['--office-teams-brand-foreground', '--office-teams-brand', 4.5],
	'link on the document page': ['--office-canvas-link', '--office-paper', 4.5],
	'focus ring on background': ['--office-ring', '--office-background', 3],
	'focus ring on surface': ['--office-ring', '--office-surface', 3],
	'accent indicator on background': ['--office-accent', '--office-background', 3],
	// Form control boundaries; `--office-border` is left to decorative dividers, which 1.4.11 exempts.
	'control border on background': ['--office-control-border', '--office-background', 3],
} as const satisfies Record<string, readonly [string, string, number]>;
type Pair = keyof typeof PAIRS;

/** Defaults the elements fall back to when the suite does not set a token on the page root. */
const FALLBACK: Record<string, string> = {
	'--office-paper': '#ffffff',
	'--office-canvas-link': '#0563c1',
	'--office-control-border': '#8a8886',
	'--office-teams-brand': 'var(--office-accent)',
	'--office-teams-brand-foreground': '#ffffff',
};

type Rgb = [number, number, number];
const luminance = ([r, g, b]: Rgb) => {
	const linear = (v: number) =>
		v / 255 <= 0.04045 ? v / 255 / 12.92 : ((v / 255 + 0.055) / 1.055) ** 2.4;
	return 0.2126 * linear(r) + 0.7152 * linear(g) + 0.0722 * linear(b);
};
const ratio = (a: Rgb, b: Rgb) => {
	const [high, low] = [luminance(a), luminance(b)].sort((x, y) => y - x) as [number, number];
	return (high + 0.05) / (low + 0.05);
};

const TOKENS = [...new Set(Object.values(PAIRS).flatMap(([fg, bg]) => [fg, bg]))];
const resolved = new Map<string, Record<string, Rgb>>();

test.describe('token contrast', () => {
	// In one worker, in order, so the colours are read once; a failure does not skip the rest.
	test.describe.configure({ mode: 'default' });

	// One page per theme and accent, read once; every pair below is arithmetic on these colours.
	test.beforeAll(async ({ browser }) => {
		for (const mode of ['light', 'dark'] as Mode[])
			for (const accent of ACCENTS) {
				const context = await browser.newContext({
					baseURL: test.info().project.use.baseURL ?? '',
				});
				await context.route(
					(url) =>
						/^https?:$/u.test(url.protocol) && !['127.0.0.1', 'localhost'].includes(url.hostname),
					(route) => route.abort(),
				);
				await context.addInitScript(
					([m, a]) => {
						(globalThis as { ooxmlSkipStart?: boolean }).ooxmlSkipStart = true;
						localStorage.setItem('ooxml-suite-theme', JSON.stringify({ mode: m, accent: a }));
					},
					[mode, accent] as const,
				);
				const page = await context.newPage();
				await openHome(page);
				const colours = await page.evaluate(
					({ tokens, fallback }) => {
						const probe = document.body.appendChild(document.createElement('span'));
						const out: Record<string, [number, number, number]> = {};
						for (const token of tokens) {
							probe.style.color = `var(${token}${fallback[token] ? `, ${fallback[token]}` : ''})`;
							const value = getComputedStyle(probe).color;
							const parts = value.match(/[\d.]+/gu)?.map(Number) ?? [];
							// rgb(r g b) or color(srgb r g b) with channels from 0 to 1.
							out[token] = (
								value.startsWith('color(')
									? parts.slice(0, 3).map((v) => Math.round(v * 255))
									: parts.slice(0, 3)
							) as [number, number, number];
						}
						probe.remove();
						return out;
					},
					{ tokens: TOKENS, fallback: FALLBACK },
				);
				resolved.set(`${mode}/${accent}`, colours);
				await context.close();
			}
	});

	for (const mode of ['light', 'dark'] as Mode[])
		for (const pair of Object.keys(PAIRS) as Pair[]) {
			const [fg, bg, needed] = PAIRS[pair];
			test(`${mode}: ${pair} reaches ${needed}:1 with every accent`, async () => {
				await gate(`contrast/${mode}/${pair}`, () => {
					const failing = ACCENTS.map((accent: Accent) => {
						const colours = resolved.get(`${mode}/${accent}`);
						if (!colours) throw new Error(`no colours for ${mode}/${accent}`);
						const [a, b] = [colours[fg], colours[bg]] as [Rgb, Rgb];
						return { accent, value: ratio(a, b), a, b };
					})
						.filter(({ value }) => value < needed)
						.map(
							({ accent, value, a, b }) => `${accent}: ${value.toFixed(2)} (rgb ${a} on rgb ${b})`,
						);
					expect(failing, `${fg} on ${bg}`).toEqual([]);
				});
			});
		}
});
