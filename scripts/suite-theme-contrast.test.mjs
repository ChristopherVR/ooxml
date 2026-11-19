// The suite's theme meets WCAG 2.x AA contrast for every accent in light and dark mode: accent
// text at 4.5:1 on every surface and selected fill it sits on, text on accent fills at 4.5:1, and
// focus rings, control boundaries and filled controls at 3:1. Colours come from the suite's own
// tables (theme-tokens.js, site/themes.css); the maths is ooxml-core/color.
import assert from 'node:assert/strict';
import { readdirSync, readFileSync } from 'node:fs';
import { test } from 'node:test';
import {
	accents,
	editorColors,
	officeThemeValues,
	resolveTheme,
} from '../apps/office-suite/src/theme-tokens.js';
import {
	contrastRatio,
	WCAG_NON_TEXT_CONTRAST,
	WCAG_TEXT_CONTRAST,
} from '../src/core/color/contrast.ts';

/** The suite shell palette per mode, read from site/themes.css. */
function sitePalette(mode) {
	const css = readFileSync(new URL('../site/themes.css', import.meta.url), 'utf8');
	const block = css.match(
		mode === 'dark' ? /:root\[data-theme='dark'\] \{([^}]*)\}/ : /:root \{([^}]*)\}/,
	)[1];
	return Object.fromEntries(
		[...block.matchAll(/--([\w-]+):\s*(#[0-9a-f]{3,6});/gi)].map((m) => [m[1], m[2]]),
	);
}

const MODES = ['light', 'dark'];
const CASES = Object.keys(accents).flatMap((accent) => MODES.map((mode) => [accent, mode]));

/** Every [label, foreground, background, minimum] pair a theme must meet. */
function pairs(accent, mode) {
	const theme = resolveTheme(mode, accent);
	const c = editorColors(theme);
	const office = officeThemeValues(theme);
	const site = sitePalette(mode);
	const editorSurfaces = {
		background: c.background,
		surface: c.secondary,
		muted: c.muted,
		selected: c.accent,
	};
	const siteSurfaces = {
		canvas: site.canvas,
		surface: site.surface,
		rail: site.rail,
		'surface-alt': site['surface-alt'],
		sunken: site.sunken,
		selected: site.selected,
	};
	const list = [];
	for (const [name, surface] of Object.entries(editorSurfaces)) {
		list.push([`accent ink on editor ${name}`, office['accent-ink'], surface, WCAG_TEXT_CONTRAST]);
		list.push([
			`Teams brand ink on ${name}`,
			office['teams-brand-ink'],
			surface,
			WCAG_TEXT_CONTRAST,
		]);
		list.push([`foreground on editor ${name}`, c.foreground, surface, WCAG_TEXT_CONTRAST]);
		list.push([
			`control border on editor ${name}`,
			office['control-border'],
			surface,
			WCAG_NON_TEXT_CONTRAST,
		]);
		list.push([`focus ring on editor ${name}`, office.ring, surface, WCAG_NON_TEXT_CONTRAST]);
	}
	list.push([
		'muted text on editor background',
		c.mutedForeground,
		c.background,
		WCAG_TEXT_CONTRAST,
	]);
	list.push([
		'checked control on editor background',
		c.primary,
		c.background,
		WCAG_NON_TEXT_CONTRAST,
	]);
	list.push(['text on accent', office['accent-foreground'], office.accent, WCAG_TEXT_CONTRAST]);
	list.push([
		'text on Teams brand',
		office['teams-brand-foreground'],
		office['teams-brand'],
		WCAG_TEXT_CONTRAST,
	]);
	for (const [name, surface] of Object.entries(siteSurfaces)) {
		list.push([`suite accent ink on ${name}`, theme.accentInk, surface, WCAG_TEXT_CONTRAST]);
		list.push([`suite focus ring on ${name}`, theme.accent, surface, WCAG_NON_TEXT_CONTRAST]);
		list.push([
			`suite control boundary on ${name}`,
			site['control-line'],
			surface,
			WCAG_NON_TEXT_CONTRAST,
		]);
	}
	// Surfaces named after where they appear, so a failure names the product surface.
	list.push([
		'maturity tag on the current rail item',
		theme.accentInk,
		site.selected,
		WCAG_TEXT_CONTRAST,
	]);
	list.push([
		'Visio current page tab on the page bar',
		office['accent-ink'],
		c.secondary,
		WCAG_TEXT_CONTRAST,
	]);
	list.push([
		'Excel selected header on its fill',
		office['accent-ink'],
		c.accent,
		WCAG_TEXT_CONTRAST,
	]);
	list.push(['ribbon tab on the ribbon', office['accent-ink'], c.background, WCAG_TEXT_CONTRAST]);
	list.push(['suite text on accent', site['accent-fg'], theme.accent, WCAG_TEXT_CONTRAST]);
	return list;
}

for (const [accent, mode] of CASES) {
	test(`${accent} accent in ${mode} mode meets AA contrast`, () => {
		const failures = pairs(accent, mode)
			.map(([label, fg, bg, min]) => [label, fg, bg, min, contrastRatio(fg, bg)])
			.filter(([, , , min, ratio]) => !(ratio >= min))
			.map(
				([label, fg, bg, min, ratio]) => `${label}: ${fg} on ${bg} ${ratio?.toFixed(2)} < ${min}`,
			);
		assert.deepEqual(failures, []);
	});
}

test('suite text takes the accent ink, never the raw accent', () => {
	// The raw accent is a fill colour; as text it falls under 4.5:1 on dark and selected surfaces.
	const site = new URL('../site/', import.meta.url);
	const offenders = readdirSync(site)
		.filter((name) => name.endsWith('.css'))
		.flatMap((name) =>
			[...readFileSync(new URL(name, site), 'utf8').matchAll(/(?:^|[;{\s])color:\s*([^;]+);/g)]
				.filter(([, value]) =>
					/var\(--(?:blue|accent)\)|var\(--tone,\s*var\(--blue\)\)/.test(value),
				)
				.map(([whole]) => `${name}: ${whole.trim()}`),
		);
	assert.deepEqual(offenders, []);
	// The maturity tag's tone defaults to the ink, so the current rail item's tag reads at 4.5:1.
	assert.match(readFileSync(new URL('styles.css', site), 'utf8'), /--tone: var\(--accent-ink\);/);
});

test('the suite leaves hover tokens to their accent-derived defaults', () => {
	for (const [accent, mode] of CASES) {
		const values = officeThemeValues(resolveTheme(mode, accent));
		assert.equal(values['accent-hover'], undefined);
		assert.equal(values['teams-brand-hover'], undefined);
	}
});
