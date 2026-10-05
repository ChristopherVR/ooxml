// Every string file of this folder merged into one table; see ../merge.ts. Listed explicitly
// (not globbed) so the package build, which does not use Vite, resolves them.
import { mergeStringModules } from '../merge.js';
import * as commands from './commands.js';
import * as grid from './grid.js';
import * as shell from './shell.js';

export const es: Readonly<Record<string, string>> = mergeStringModules({
	'./commands.ts': commands,
	'./grid.ts': grid,
	'./shell.ts': shell,
});
