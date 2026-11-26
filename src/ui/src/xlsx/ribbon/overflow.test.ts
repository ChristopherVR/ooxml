import { beforeAll, expect, it } from 'vitest';
import { registerSheetIcons } from '../commands/register-icons';
import { ribbonIcon } from './icons';
import { GROUP_ICONS } from './overflow';

beforeAll(registerSheetIcons);

it.each(Object.entries(GROUP_ICONS))(
	'draws a path for the folded %s group (%s)',
	(_group, icon) => {
		expect(ribbonIcon(document, icon).querySelector('path')?.getAttribute('d')).toMatch(/\S/u);
	},
);
