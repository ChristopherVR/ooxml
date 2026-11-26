import { expect, it } from 'vitest';
import { ribbonIcon } from './ribbon-icons';
import { GROUP_ICONS } from './ribbon-overflow';

it.each(Object.entries(GROUP_ICONS))(
	'draws a path for the folded %s group (%s)',
	(_group, icon) => {
		expect(ribbonIcon(icon).querySelector('path')?.getAttribute('d')).toMatch(/^[Mm]\s*[-.\d]/u);
	},
);
