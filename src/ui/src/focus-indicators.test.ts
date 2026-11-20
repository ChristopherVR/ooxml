// A11Y-18g: the controls whose focus style was a recoloured 1px border, or a ring drawn in the
// colour it sat on. A focus indicator must be a real change of at least 3:1 (WCAG 1.4.11), so each
// of these now draws the shared ring; the measured proof is the focus-ring diff in the a11y gates.
import { expect, it } from 'vitest';
import commandSearch from './menu/command-search.css?raw';
import ribbonTabs from './ribbon/ribbon-tabs.css?raw';
import searchField from './form/search-field.css?raw';
import { OFFICE_TOKENS } from './tokens';
import { FORMULA_BAR_CSS } from './xlsx/formula-bar/styles';

/** Everything every rule declares for `selector`, comments stripped first. */
function rule(css: string, selector: string): string {
	const found = [...css.replace(/\/\*[\s\S]*?\*\//g, '').matchAll(/([^{}]+)\{([^}]*)\}/g)]
		.filter((match) =>
			match[1]!
				.split(',')
				.map((part) => part.trim())
				.includes(selector),
		)
		.map((match) => match[2]!);
	if (!found.length) throw new Error(`no rule for ${selector}`);
	return found.join(';');
}

it('rings the Tell me box, not just its border, when focus is inside it', () => {
	const declarations = rule(commandSearch, '.field:focus-within');
	expect(declarations).toContain('outline: var(--office-focus-width) solid var(--office-ring)');
	expect(declarations).toContain('outline-offset: calc(-1 * var(--office-focus-width))');
});

it('rings a search field, not just its border', () => {
	const declarations = rule(searchField, ':host(:focus-within)');
	expect(declarations).toContain('outline: var(--office-focus-width) solid var(--office-ring)');
});

it('rings the Excel Name Box and formula bar, not just their borders', () => {
	const declarations = rule(FORMULA_BAR_CSS, '.xfb-namebox:focus-within');
	expect(declarations).toContain('outline:var(--office-focus-width, 2px) solid');
	expect(rule(FORMULA_BAR_CSS, '.xfb-input:focus-within')).toContain('outline:');
});

it('rings the ribbon File tab in a colour its accent fill contrasts with', () => {
	// `--office-ring` is a shade of the accent, which is what the File tab is filled with: the ring
	// was being drawn in the colour it sat on, so focus showed nothing at all.
	expect(rule(ribbonTabs, '.file:focus-visible')).toContain('solid var(--office-ring-on-accent)');
	expect(OFFICE_TOKENS['--office-ring-on-accent']).toContain('--office-accent-foreground');
	expect(rule(ribbonTabs, "[role='tab']:focus-visible")).toContain('solid var(--office-ring)');
});
