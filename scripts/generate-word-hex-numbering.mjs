import { hexNumberingFixture } from '../src/core/docx/test-support/hex-numbering-fixture.ts';

const output = process.argv[2];
if (!output) throw new Error('Usage: bun scripts/generate-word-hex-numbering.mjs <output.docx>');
await Bun.write(output, await hexNumberingFixture());
