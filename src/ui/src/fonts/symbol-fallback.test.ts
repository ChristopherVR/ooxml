import { afterEach, expect, it } from 'vitest';
import { SYMBOL_FALLBACK_CSS, withSymbolFallback } from './symbol-fallback';

afterEach(() => document.getElementById('ooxml-symbol-fallback')?.remove());

it('keeps Georgia first and supplies symbols before the generic fallback, offline', () => {
	expect(withSymbolFallback('"Georgia", serif')).toBe('"Georgia", "OOXML Noto Symbols", serif');
	expect(document.getElementById('ooxml-symbol-fallback')?.textContent).toBe(SYMBOL_FALLBACK_CSS);
	expect(SYMBOL_FALLBACK_CSS).toContain('data:font/woff2;base64,');
	expect(SYMBOL_FALLBACK_CSS).not.toMatch(/https?:/u);
	withSymbolFallback('"Georgia", serif');
	expect(document.querySelectorAll('#ooxml-symbol-fallback')).toHaveLength(1);
});

it('does not duplicate the fallback and works with only a generic family', () => {
	const family = withSymbolFallback('sans-serif');
	expect(family).toBe('"OOXML Noto Symbols", sans-serif');
	expect(withSymbolFallback(family)).toBe(family);
});
