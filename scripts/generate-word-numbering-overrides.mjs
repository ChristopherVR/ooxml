// Run with Bun; this shares the exact portable fixture generator exercised by core tests.
import { readFile, mkdir } from 'node:fs/promises';
import { resolve } from 'node:path';
import { overrideRestartFixture } from '../src/core/docx/test-support/override-restart-fixture.ts';
import { numberingStartFixture } from '../src/core/docx/test-support/numbering-start-fixture.ts';
import { authoredNumberingStartFixture } from '../src/core/docx/test-support/authored-numbering-start-fixture.ts';

const output = process.argv[2];
const kind = process.argv[3] ?? 'restart';
if (!['restart', 'start', 'edit'].includes(kind))
	throw new Error('Fixture kind must be restart, start or edit');
if (!output)
	throw new Error('Usage: bun scripts/generate-word-numbering-overrides.mjs <output-directory>');
const reference = JSON.parse(
	await readFile(
		new URL(
			`../src/core/docx/__fixtures__/numbering-override-${kind}/native-reference.json`,
			import.meta.url,
		),
		'utf8',
	),
);
await mkdir(output, { recursive: true });
for (const item of reference.cases) {
	await Bun.write(
		resolve(output, `${item.name}.docx`),
		kind === 'edit'
			? await authoredNumberingStartFixture(item)
			: kind === 'start'
				? await numberingStartFixture(item)
				: await overrideRestartFixture(item.abstractRestart ?? undefined, item.overrideRestart),
	);
}
