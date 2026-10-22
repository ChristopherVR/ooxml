import { readFileSync } from 'node:fs';
import path from 'node:path';

import { describe, expect, it } from 'vitest';

const stylesDir = import.meta.dirname;

describe('styles.css', () => {
	it('replaces the global Tailwind preflight with the viewer-scoped one', () => {
		const css = readFileSync(path.join(stylesDir, 'pptx-viewer.css'), 'utf8');
		expect(css).not.toMatch(/@import\s+'tailwindcss'\s*;/u);
		expect(css).toContain("@import 'tailwindcss/theme.css' layer(theme);");
		expect(css).toContain(
			"@import '../../../../../../src/ui/src/pptx/theme/viewer-preflight.css' layer(base);",
		);
		expect(css).toContain("@import 'tailwindcss/utilities.css' layer(utilities);");
	});
});
