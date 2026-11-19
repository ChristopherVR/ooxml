import { describe, expect, it } from 'vitest';

import { LEFT_OUT_MESSAGE, UPDATE_REFUSAL_MESSAGES, UpdateRefusedError } from '../collab/refusal';
import { collabStrings, translateCollabText } from './collab-strings';

const english = [...Object.values(UPDATE_REFUSAL_MESSAGES), LEFT_OUT_MESSAGE];

describe('collabStrings', () => {
	it('translates every refusal and the left-out notice in every locale', () => {
		for (const [locale, table] of Object.entries(collabStrings)) {
			expect(Object.keys(table).sort()).toEqual([...english].sort());
			for (const text of english)
				if (locale === 'en') expect(table[text]).toBe(text);
				else expect(table[text]).not.toBe(text);
		}
	});

	it("translates a refusal's message, falling back to the base language and to the text", () => {
		const { message } = new UpdateRefusedError('server-full', 'document too large');
		expect(translateCollabText(message, 'de')).toBe(collabStrings.de[message]);
		expect(translateCollabText(message, 'fr-CA')).toBe(collabStrings.fr[message]);
		expect(translateCollabText(message, 'zh-CN')).toBe(collabStrings['zh-CN'][message]);
		expect(translateCollabText('Something else', 'de')).toBe('Something else');
	});

	it('reads any tag of a shipped language, as the product UIs resolve their locale', () => {
		const message = UPDATE_REFUSAL_MESSAGES['too-large'];
		const chinese = collabStrings['zh-CN'][message];
		for (const tag of ['zh', 'zh-Hans', 'zh-Hans-CN', 'zh_CN', 'ZH-cn', 'zh-SG']) {
			expect(translateCollabText(message, tag)).toBe(chinese);
		}
		expect(translateCollabText(message, 'de-AT')).toBe(collabStrings.de[message]);
		expect(translateCollabText(message, 'ja-JP')).toBe(collabStrings.ja[message]);
		expect(translateCollabText(message, 'es-419')).toBe(collabStrings.es[message]);
		// No product ships Traditional Chinese: English, not Simplified.
		expect(translateCollabText(message, 'zh-TW')).toBe(message);
		expect(translateCollabText(message, 'zh-Hant')).toBe(message);
		expect(translateCollabText(message, undefined)).toBe(message);
		expect(translateCollabText(message, 'pt-BR')).toBe(message);
	});
});
