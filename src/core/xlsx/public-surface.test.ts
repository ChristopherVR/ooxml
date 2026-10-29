import { expect, test } from 'vitest';

import * as root from './index';
import * as load from './load/index';

// The runtime export names of the entries that `@christophervr/xlsx-core` re-exports. Adding or
// removing an export is a visible diff here; docs/stability.md says which names are supported
// API. Types are erased and are not listed.
const names = (module: object) => `${JSON.stringify(Object.keys(module).sort(), null, '\t')}\n`;

test('the xlsx root entry exports a stable set of names', async () => {
	await expect(names(root)).toMatchFileSnapshot('./__snapshots__/public-surface.json');
});

test('the xlsx load entry exports a stable set of names', async () => {
	await expect(names(load)).toMatchFileSnapshot('./__snapshots__/public-surface-load.json');
});
