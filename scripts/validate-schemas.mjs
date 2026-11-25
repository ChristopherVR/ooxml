// Generates DOCX, XLSX, PPTX and VSDX fixtures from the core writers, validates the OOXML
// packages with Microsoft's OpenXmlValidator (.NET) and the VSDX packages with the Visio
// structural package validator, and cleans up the temporary output directory.
// Usage: node scripts/validate-schemas.mjs [--version Microsoft365]
import { spawnSync } from 'node:child_process';
import { mkdtempSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';

function run(command, args) {
	const result = spawnSync(command, args, { stdio: 'inherit' });
	if (result.error) throw result.error;
	if (result.status !== 0) process.exit(result.status ?? 1);
}

const extraArgs = process.argv.slice(2);
const outDir = mkdtempSync(join(tmpdir(), 'ooxml-schema-validation-'));
try {
	run(process.execPath, ['scripts/generate-schema-fixtures.ts', outDir]);
	run('dotnet', ['run', '--project', 'scripts/ooxml-validate', '--', ...extraArgs, outDir]);
	run(process.execPath, ['scripts/validate-vsdx.ts', outDir]);
} finally {
	rmSync(outDir, { recursive: true, force: true, maxRetries: 5, retryDelay: 100 });
}
