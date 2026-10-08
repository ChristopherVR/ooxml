import { cp, mkdir, readdir, readFile, rm } from 'node:fs/promises';
import { join, relative } from 'node:path';

const sourceDir = join(process.cwd(), process.argv[2] ?? '.types');
const outputDir = join(process.cwd(), 'dist');

const declarationFiles = [];

// An export list may only name local bindings. The declaration bundler has
// emitted `ns.Name as Name` there for a re-export reached through `export *`
// (issue #33), which every consumer's `tsc` rejects as a syntax error even
// with `skipLibCheck`, so fail the build instead of publishing it.
const EXPORT_LIST = /export\s*(?:type\s*)?\{([^}]*)\}/g;
const DOTTED_NAME = /(?:^|,)\s*(?:type\s+)?[\w$]+\.[\w$.]+/;

async function checkExportLists(file) {
	const text = await readFile(file, 'utf8');
	for (const match of text.matchAll(EXPORT_LIST)) {
		const dotted = DOTTED_NAME.exec(match[1]);
		if (dotted) {
			throw new Error(
				`${relative(process.cwd(), file)}: dotted name in an export list: ${dotted[0].replace(/^,?\s*/, '')}`,
			);
		}
	}
}

async function copyDeclarations(directory) {
	for (const entry of await readdir(directory, { withFileTypes: true })) {
		const sourcePath = join(directory, entry.name);
		if (entry.isDirectory()) {
			await copyDeclarations(sourcePath);
			continue;
		}

		if (!entry.name.endsWith('.d.ts') && !entry.name.endsWith('.d.ts.map')) {
			continue;
		}

		const destination = join(outputDir, relative(sourceDir, sourcePath));
		await mkdir(join(destination, '..'), { recursive: true });
		await cp(sourcePath, destination);
		if (!entry.name.endsWith('.map')) declarationFiles.push(destination);
	}
}

await copyDeclarations(sourceDir);
for (const file of declarationFiles) await checkExportLists(file);
await rm(sourceDir, { recursive: true, force: true });
