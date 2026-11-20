// @vitest-environment node
// The source half of the element conformance gate (elements.test.ts checks the live classes):
// what an element's module must look like, read from every non-test source file.
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { describe, expect, it } from 'vitest';
import { sourcesMatching } from './elements';

const text = (module: string): string =>
	readFileSync(resolve(import.meta.dirname, '..', module), 'utf8');

describe('element sources', () => {
	it('define elements only through the registry, so every class is stamped and adopts early properties', () => {
		const offenders = sourcesMatching(/\b(customElements|registry|target)\.define\(/).filter(
			(module) => module !== 'registry.ts',
		);
		expect(offenders).toEqual([]);
	});

	it('take shadow styles from a sibling ?raw stylesheet, never from CSS written in the module', () => {
		const offenders = sourcesMatching(/\bstatic (override )?(get )?styles\b/).filter(
			(module) => !/from '\.{1,2}\/[\w./-]+\.css\?raw'/.test(text(module)),
		);
		expect(offenders).toEqual([]);
	});
});
