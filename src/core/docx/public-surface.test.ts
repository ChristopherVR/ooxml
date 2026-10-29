import { expect, test } from 'vitest';

import * as embedded from './embedded';
import * as root from './index';

// The runtime export names of the entries that `docx-core` re-exports. Adding or removing an
// export is a visible diff here; docs/stability.md says which names are supported API. Types are
// erased and are not listed.
const names = (module: object) => `${JSON.stringify(Object.keys(module).sort(), null, '\t')}\n`;

test('the docx root entry exports a stable set of names', async () => {
	await expect(names(root)).toMatchFileSnapshot('./__snapshots__/public-surface.json');
});

test('the docx embedded entry exports a stable set of names', async () => {
	await expect(names(embedded)).toMatchFileSnapshot('./__snapshots__/public-surface-embedded.json');
});
