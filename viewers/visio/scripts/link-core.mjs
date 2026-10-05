import { lstatSync, mkdirSync, readFileSync, readlinkSync, symlinkSync, unlinkSync } from 'node:fs';
import { dirname, relative, resolve } from 'node:path';
import { coreDirectory, viewerRoot } from './core-paths.mjs';

const metadata = JSON.parse(readFileSync(resolve(coreDirectory, 'package.json'), 'utf8'));
if (metadata.name !== 'ooxml-core' || !metadata.exports?.['./visio'])
	throw new Error('Selected checkout does not export ooxml-core/visio.');
const dependency = resolve(viewerRoot, 'node_modules/ooxml-core');
const installed = lstatSync(dependency, { throwIfNoEntry: false });
if (installed && !installed.isSymbolicLink())
	throw new Error('Installed ooxml-core is not a symlink. Nothing was overwritten.');
if (installed && resolve(dirname(dependency), readlinkSync(dependency)) === coreDirectory) {
	console.log(`Installed ooxml-core already points at ${coreDirectory}.`);
} else {
	mkdirSync(dirname(dependency), { recursive: true });
	if (installed) unlinkSync(dependency);
	symlinkSync(relative(dirname(dependency), coreDirectory), dependency, 'junction');
	console.log(
		`Linked installed ooxml-core to ${coreDirectory}. Package manifest and lock are unchanged.`,
	);
}
