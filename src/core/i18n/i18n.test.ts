import { describe, expect, it } from 'vitest';
import {
	SUPPORTED_OFFICE_LOCALES,
	createTranslator,
	mergeMessageModules,
	missingMessageKeys,
	normalizeLocale,
	normalizeOfficeLocale,
} from './index';

describe('normalizeLocale', () => {
	it.each([
		['fr', 'fr'],
		['FR_ca', 'fr'],
		['de-AT', 'de'],
		['es-MX', 'es'],
		['zh', 'zh-CN'],
		['zh-CN', 'zh-CN'],
		['zh_cn', 'zh-CN'],
		['zh-Hans-CN', 'zh-CN'],
		['zh-SG', 'zh-CN'],
		['zh-TW', 'en'],
		['zh-Hant', 'en'],
		['ja', 'en'],
		['', 'en'],
		[null, 'en'],
		[undefined, 'en'],
	])('%s -> %s', (input, expected) => {
		expect(normalizeOfficeLocale(input)).toBe(expected);
	});
	it('uses the supplied fallback and set', () => {
		expect(normalizeLocale('it', ['de', 'fr'], 'fr')).toBe('fr');
		expect(SUPPORTED_OFFICE_LOCALES).toHaveLength(5);
	});
});

describe('dictionary', () => {
	it('merges string tables in path order and skips non-string objects', () => {
		const merged = mergeMessageModules({
			'./b.ts': { b: { x: 'B', y: 'B2' } },
			'./a.ts': { default: { x: 'A' }, count: 3, mixed: { z: 1 } },
		});
		expect(merged).toEqual({ x: 'B', y: 'B2' });
		expect(missingMessageKeys({ x: '', q: '' }, merged)).toEqual(['q']);
	});
});

describe('createTranslator', () => {
	it('falls back to English, then the key, and fills placeholders', () => {
		const t = createTranslator({ a: 'Hallo {n}' }, [{ b: 'Bee' }]);
		expect(t('a', { n: 2 })).toBe('Hallo 2');
		expect(t('b')).toBe('Bee');
		expect(t('c {n}', {})).toBe('c {n}');
		expect(t('a', {})).toBe('Hallo {n}');
	});
	it('supports double-brace placeholders and a missing formatter', () => {
		const t = createTranslator(undefined, [], {
			placeholder: 'double',
			missing: (key) => key.toUpperCase(),
		});
		expect(t('x {{ n }}', { n: 1 })).toBe('X {{ N }}');
	});
});
