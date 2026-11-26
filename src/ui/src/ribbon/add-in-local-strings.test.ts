import { expect, it } from 'vitest';
import { localAddInText } from './add-in-local-strings';

it.each(['fr', 'de', 'es', 'zh-CN', 'ja', 'ko'])(
	'translates all add-in detail labels in %s',
	(locale) => {
		const text = localAddInText(locale, () => '');
		for (const label of ['Requirements', 'Icons', 'Open', 'Close'])
			expect(text(label)).not.toBe('');
	},
);
