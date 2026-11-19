import { describe, expect, it } from 'vitest';
import { UpdateRefusedError } from 'ooxml-core/collab';
import { collabStrings } from 'ooxml-core/i18n';
import { localizeCoreText } from './core-text';
import { teamsI18n } from './translator';

describe('localizeCoreText', () => {
	it('says why sharing stopped in the display locale', () => {
		const { message } = new UpdateRefusedError('server-full', 'document too large');
		expect(localizeCoreText(message, teamsI18n('de'))).toBe(collabStrings.de[message]);
		expect(localizeCoreText(message, teamsI18n('zh-CN'))).toBe(collabStrings['zh-CN'][message]);
		expect(localizeCoreText(message, teamsI18n('en'))).toBe(message);
	});
});
