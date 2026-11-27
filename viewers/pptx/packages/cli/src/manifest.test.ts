import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

import { describe, expect, it } from 'vitest';

/**
 * Importing the package root behaves like `pptx-react-viewer`, so this package must ask for
 * exactly the peers that package asks for (see `./index.ts`).
 */
type Manifest = { peerDependencies: Record<string, string> };
const readManifest = (path: string): Manifest =>
	JSON.parse(readFileSync(resolve(__dirname, path), 'utf8')) as Manifest;

describe('package manifest', () => {
	it('declares every peer pptx-react-viewer declares, with the same range', () => {
		const own = readManifest('../package.json').peerDependencies;
		const react = readManifest('../../react/package.json').peerDependencies;
		expect(own).toStrictEqual(react);
	});
});
