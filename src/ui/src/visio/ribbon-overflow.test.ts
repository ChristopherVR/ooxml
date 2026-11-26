import { beforeAll, expect, it } from 'vitest';
import { getIcon } from '../icons';
import { registerViewerControls } from './office-ui';
import { GROUP_ICONS } from './ribbon-overflow';

beforeAll(() => registerViewerControls());

it.each(Object.entries(GROUP_ICONS))(
	'draws a path for the folded %s group (%s)',
	(_group, icon) => {
		expect(getIcon(icon)?.d).toMatch(/\S/u);
	},
);
