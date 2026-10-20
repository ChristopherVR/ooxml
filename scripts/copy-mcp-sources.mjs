// Copy an MCP package into a clean consumer without Bun's workspace dependency links.
import { cpSync } from 'node:fs';
import { relative, resolve, sep } from 'node:path';

const [source, target] = process.argv.slice(2);
if (!source || !target) throw new Error('Expected source and target directories');
const root = resolve(source);
cpSync(root, resolve(target), {
	recursive: true,
	filter: (file) => !relative(root, file).split(sep).includes('node_modules'),
});
