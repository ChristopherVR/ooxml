import { describe, expect, it } from 'vitest';

import { tokenizeHtml } from './html-tokens';

describe('tokenizeHtml past its token limit', () => {
	it('returns the rest as one text token with complete tags removed', () => {
		const tokens = tokenizeHtml('<b>one</b> a < b <i>two</i> <unclosed', 1);
		expect(tokens.at(-1)).toEqual({ kind: 'text', text: 'one a < b two <unclosed' });
	});

	it('stays linear on many unclosed angle brackets', () => {
		const started = performance.now();
		tokenizeHtml('<b>x</b>' + '<'.repeat(200_000), 1);
		expect(performance.now() - started).toBeLessThan(2000);
	});
});
