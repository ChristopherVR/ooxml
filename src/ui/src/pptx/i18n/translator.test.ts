import { createTranslator as shared } from 'ooxml-core/i18n';
import { describe, expect, it } from 'vitest';
import { keyToLabel, translationsEn } from './translations-en';
import { createTranslator } from './translator';

describe('pptx translator over ooxml-core/i18n', () => {
	const messages = { de: { 'pptx.a': 'Hallo {{ n }}' }, en: { 'pptx.b': 'Bee {{n}}' } };
	it('equals the shared translator with the pptx fallbacks', () => {
		const wrapped = createTranslator('de', messages);
		const direct = shared(messages.de, [messages.en, translationsEn], {
			placeholder: 'double',
			missing: keyToLabel,
		});
		for (const key of ['pptx.a', 'pptx.b', 'pptx.unknown.key_name']) {
			expect(wrapped(key, { n: 3 })).toBe(direct(key, { n: 3 }));
		}
		expect(wrapped('pptx.a', { n: 3 })).toBe('Hallo 3');
		expect(wrapped('pptx.b', { n: 4 })).toBe('Bee 4');
		expect(wrapped('pptx.a')).toBe('Hallo {{ n }}');
	});
});
