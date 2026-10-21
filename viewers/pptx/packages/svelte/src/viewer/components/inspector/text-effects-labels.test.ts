import { describe, expect, it } from 'vitest';

import { translate } from '../../../i18n/translator';
import type { Translator } from '../../../i18n/translator';
import { bevelPresetLabel, materialPresetLabel } from './text-effects-labels';

const en: Translator = (key, params) => translate('en', key, params);

describe('text-effects labels', () => {
	it('labels the 3D material and bevel presets', () => {
		expect(materialPresetLabel('warmMatte', en)).toBe('Warm Matte');
		expect(materialPresetLabel('softEdge', en)).toBe('Soft Edge');
		expect(bevelPresetLabel('relaxedInset', en)).toBe('Relaxed Inset');
		expect(bevelPresetLabel('coolSlant', en)).toBe('Cool Slant');
	});

	it('falls back to the raw token rather than blanking an unknown preset', () => {
		// A deck may carry a value newer than the catalogue; an empty option
		// would read as a broken control instead of an untranslated one.
		expect(materialPresetLabel('somethingNew', en)).toBe('somethingNew');
		expect(bevelPresetLabel('somethingNew', en)).toBe('somethingNew');
	});
});
