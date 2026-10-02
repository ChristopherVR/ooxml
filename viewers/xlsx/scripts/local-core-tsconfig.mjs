// Writes the git-ignored tsconfig.local.json, which type-checks against the TypeScript sources of
// a local `ooxml` checkout (OOXML_CORE_SRC, for example `../ooxml`) instead of the published
// ooxml-core and ooxml-ui. Use it with `bun run typecheck:local`; CI never does.
import { readFileSync, writeFileSync, existsSync } from 'node:fs';
import { relative, resolve, sep } from 'node:path';

const checkout = process.env.OOXML_CORE_SRC;
if (!checkout) throw new Error('Set OOXML_CORE_SRC to your ooxml checkout, e.g. ../ooxml');
const root = resolve(import.meta.dirname, '..');
const base = resolve(root, checkout);
if (!existsSync(resolve(base, 'src/xlsx/index.ts')))
	throw new Error(`${base} has no src/xlsx/index.ts`);
const rel = (path) => relative(root, resolve(base, path)).split(sep).join('/');
const tsconfig = JSON.parse(readFileSync(resolve(root, 'tsconfig.json'), 'utf8'));
const paths = {
	...tsconfig.compilerOptions.paths,
	'ooxml-core/xlsx/load': [rel('src/xlsx/load/index.ts')],
	'ooxml-core/*': [rel('src/*/index.ts')],
	'ooxml-core': [rel('src/index.ts')],
	'ooxml-ui/*': [rel('packages/ui/src/*.ts')],
	'ooxml-ui': [rel('packages/ui/src/index.ts')],
};
writeFileSync(
	resolve(root, 'tsconfig.local.json'),
	`${JSON.stringify({ extends: './tsconfig.json', compilerOptions: { paths } }, null, '\t')}\n`,
);
console.log(`tsconfig.local.json -> ${base}`);
