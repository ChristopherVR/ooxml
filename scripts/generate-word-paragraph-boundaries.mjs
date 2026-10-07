import { resolve } from 'node:path';
import { mkdir } from 'node:fs/promises';
import { paragraphBoundaryFixture } from '../src/core/docx/test-support/paragraph-boundary-fixture.ts';

const output = process.argv[2];
if (!output)
	throw new Error('Usage: bun scripts/generate-word-paragraph-boundaries.mjs <output-directory>');
await mkdir(output, { recursive: true });
for (const action of ['split', 'join'])
	await Bun.write(resolve(output, `${action}.docx`), await paragraphBoundaryFixture(action));
