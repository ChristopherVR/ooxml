import { describe, expect, it } from 'vitest';
import { readFileSync } from 'node:fs';
import native from './__fixtures__/word-theme-colors.json';
import { isThemeColorToken, resolveThemeColorReference, themeByteToFraction } from './theme-color';
import type { ThemeColorReference } from './theme-model';
import { parseTheme } from './theme';

const theme = parseTheme(
	readFileSync(new URL('./__fixtures__/word-theme-native.xml', import.meta.url), 'utf8'),
);
interface NativeColor {
	themeColor: string;
	expected: string;
	themeTint?: string;
	themeShade?: string;
}
function reference(entry: NativeColor): ThemeColorReference {
	if (!isThemeColorToken(entry.themeColor)) throw new Error('Unexpected native theme token.');
	const result: ThemeColorReference = { token: entry.themeColor };
	const tint = themeByteToFraction(entry.themeTint);
	const shade = themeByteToFraction(entry.themeShade);
	if (tint !== undefined) result.tint = tint;
	if (shade !== undefined) result.shade = shade;
	return result;
}

describe('native Word theme colours', () => {
	it.each(native.backgrounds)('matches a native page background: $expected', (entry) => {
		expect(resolveThemeColorReference(reference(entry), theme)).toBe(`#${entry.expected}`);
	});

	it('keeps native text tint/shade errors within one channel level', () => {
		// HSL is correct; Word's final quantization still differs in some cases by one level.
		// Keep that remaining parity gap explicit instead of accepting the old linear-light mix.
		expect(native.colors).toHaveLength(90);
		for (const entry of native.colors) {
			const actual = resolveThemeColorReference(reference(entry), theme)!;
			for (const offset of [0, 2, 4]) {
				const expected = Number.parseInt(entry.expected.slice(offset, offset + 2), 16);
				const received = Number.parseInt(actual.slice(offset + 1, offset + 3), 16);
				expect(Math.abs(received - expected), JSON.stringify(entry)).toBeLessThanOrEqual(1);
			}
		}
	});

	it('gives tint precedence over shade as Word does', () => {
		expect(
			resolveThemeColorReference({ token: 'accent2', tint: 127 / 255, shade: 0.25 }, theme),
		).toBe('#F6BE98');
	});
});
