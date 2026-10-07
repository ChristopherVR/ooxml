import { mkdir } from 'node:fs/promises';
import { resolve } from 'node:path';
import { fieldLockFixture } from '../src/core/docx/test-support/field-lock-fixture.ts';

const output = process.argv[2];
if (!output) throw new Error('Usage: bun scripts/generate-word-field-locks.mjs <directory>');
await mkdir(output, { recursive: true });
for (const kind of ['simple', 'complex'])
	await Bun.write(resolve(output, `${kind}.docx`), await fieldLockFixture(kind));
