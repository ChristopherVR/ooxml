import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { describe, expect, it } from 'vitest';
import { canvasAndRibbonStyles, visioThemeAliases } from './index';
import { viewerStyles } from '../styles';

const viewerElementSource = readFileSync(join(import.meta.dirname, '../viewer-element.ts'), 'utf8');

describe('visio theme', () => {
	it('references --vv-* colours only in the alias block', () => {
		const rest = `${viewerStyles}\n${canvasAndRibbonStyles}`;
		// --vv-inch is the canvas scale, not a theme token.
		const stray = rest.match(/--vv-(?!inch\b)[a-z-]+/g);
		expect(stray).toBeNull();
		expect(visioThemeAliases).toContain('--vv-ink');
	});

	it('derives every alias from the shared --office-* tokens', () => {
		expect(visioThemeAliases).toContain('var(--office-foreground,');
		expect(visioThemeAliases).toContain('var(--office-accent,');
		expect(visioThemeAliases).toContain('var(--office-background,');
	});

	it('bridges the aliases back to --office-* for the shared controls', () => {
		expect(canvasAndRibbonStyles).toContain(':host > * {');
		expect(canvasAndRibbonStyles).toContain('--office-foreground: var(--_vv-ink);');
	});

	it('installs the shared theme and the alias block from the viewer element', () => {
		expect(viewerElementSource).toContain('installOfficeUiTheme()');
		expect(viewerElementSource).toContain('${visioThemeAliases}');
	});
});
