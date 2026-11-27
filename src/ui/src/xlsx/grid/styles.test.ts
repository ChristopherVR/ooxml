import { expect, it } from 'vitest';
import { LOOK_TOKENS } from '../../look-tokens';
import { OFFICE_TOKENS } from '../../tokens';
import { GRID_CSS } from './styles';

const known = new Set([...Object.keys(OFFICE_TOKENS), ...Object.keys(LOOK_TOKENS)]);

it('reads only --office-* tokens the theme defines', () => {
	// An unknown token always falls back to its literal, so a dark theme never reaches it.
	const used = new Set(GRID_CSS.match(/--office-[a-z0-9-]+/g) ?? []);
	expect([...used].filter((name) => !known.has(name))).toEqual([]);
});

it('themes the PivotChart field buttons with the shared surface and foreground', () => {
	const rule = GRID_CSS.match(/\.xg-pivot-chart-buttons button\{[^}]*\}/)?.[0] ?? '';
	expect(rule).toContain('background:var(--office-surface');
	expect(rule).toContain('color:var(--office-foreground');
	const forced = GRID_CSS.slice(GRID_CSS.indexOf('@media (forced-colors:active)'));
	expect(forced).toMatch(/\.xg-pivot-chart-buttons button\{background:ButtonFace;color:ButtonText/);
});
