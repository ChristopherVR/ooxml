import { describe, expect, it } from 'vitest';
import { translateInkText } from './ink-strings';

describe('shared ink strings', () => {
	it('translates for any tag the products accept and fills the angle', () => {
		expect(translateInkText('Clear', 'de-AT')).toBe('Löschen');
		expect(translateInkText('Ruler', 'zh-Hans')).toBe('标尺');
		expect(translateInkText('Ruler', 'zh-TW')).toBe('Ruler');
		expect(translateInkText('{angle} degrees', 'fr', { angle: -45 })).toBe('-45 degrés');
	});

	it('has every string in every locale', () => {
		for (const locale of ['fr', 'de', 'es', 'zh-CN', 'ja', 'ko'])
			for (const text of ['Ink Equation', 'Write math here', 'Ruler', 'Recognizing...'] as const)
				expect(translateInkText(text, locale)).not.toBe(text);
	});
});
