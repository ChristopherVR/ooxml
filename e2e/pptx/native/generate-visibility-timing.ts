import { mkdir, writeFile } from 'node:fs/promises';
import { resolve } from 'node:path';

import { createVisibilityTimingDeck } from '../../../src/core/pptx/__tests__/helpers/visibility-timing-deck';

const output = resolve(process.argv[2] ?? 'visibility-timing-native');
await mkdir(output, { recursive: true });
for (const [name, delay] of [
	['zero', 0],
	['genuine', 499],
	['long', 1499],
] as const) {
	await writeFile(resolve(output, `${name}.pptx`), await createVisibilityTimingDeck(delay));
}
