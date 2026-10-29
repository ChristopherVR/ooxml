import { expect, test } from 'vitest';

import * as root from './index';

// The runtime export names of `ooxml-core/visio`. Adding or removing an export is a visible diff
// here; docs/stability.md says which names are supported API. Types are erased and are not listed.
const names = (module: object) => `${JSON.stringify(Object.keys(module).sort(), null, '\t')}\n`;

test('the visio root entry exports a stable set of names', async () => {
	await expect(names(root)).toMatchFileSnapshot('./__snapshots__/public-surface.json');
});
