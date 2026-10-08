import { normalizeLocale, SUPPORTED_OFFICE_LOCALES } from 'ooxml-core/i18n';
import { describe, expect, it } from 'vitest';
import { EDITOR_LOCALES, normalizeEditorLocale } from './localization';

describe('docx localization over ooxml-core/i18n', () => {
	it('shares the locale set and normalisation', () => {
		expect(EDITOR_LOCALES).toBe(SUPPORTED_OFFICE_LOCALES);
		for (const tag of ['de-AT', 'zh_Hans_CN', 'zh-TW', 'pt', 'FR', null]) {
			expect(normalizeEditorLocale(tag)).toBe(normalizeLocale(tag, EDITOR_LOCALES, 'en'));
		}
	});
});
