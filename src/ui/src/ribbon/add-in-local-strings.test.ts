import { expect, it } from 'vitest';
import { addInIssueLabel } from './add-in-local-dialog';
import { localAddInText } from './add-in-local-strings';

it.each(['fr', 'de', 'es', 'zh-CN', 'ja', 'ko'])(
	'translates all add-in detail labels in %s',
	(locale) => {
		const text = localAddInText(locale, () => '');
		for (const label of ['Requirements', 'Icons', 'Open', 'Close'])
			expect(text(label)).not.toBe('');
	},
);

it.each(['fr', 'de', 'es', 'zh-CN', 'ja', 'ko'])(
	'localizes known compatibility issues in %s',
	(locale) => {
		const text = localAddInText(locale, () => '');
		const prefix = text('Unsupported feature');
		for (const feature of [
			'ExecuteFunction',
			'shared runtime',
			'mobile form factor',
			'dynamic ribbon command',
			'VersionOverrides',
		]) {
			const label = addInIssueLabel({ kind: 'unsupported', feature }, text);
			expect(label).toMatch(/^\S.* · \S/u);
			expect(label.toLowerCase()).not.toContain(feature.toLowerCase());
		}
		expect(addInIssueLabel({ kind: 'unsupported', feature: 'Unknown' }, text)).toBe(prefix);
	},
);
