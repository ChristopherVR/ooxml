// @vitest-environment jsdom
import { describe, expect, it } from 'vitest';
import { labelNameMismatches } from '../test-support/label-in-name';
import { EDITOR_LOCALES } from './localization';
import { createRibbon } from './ribbon';

describe('Word ribbon names (WCAG 2.5.3)', () => {
	for (const locale of EDITOR_LOCALES)
		it(`names every captioned command with its caption in ${locale}`, () => {
			const ribbon = createRibbon(locale);
			document.body.append(ribbon);
			try {
				expect(labelNameMismatches(ribbon)).toEqual([]);
			} finally {
				ribbon.remove();
			}
		});
});
