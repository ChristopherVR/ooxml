// Every string file of this folder merged into one table; see ../merge.ts. Listed explicitly
// (not globbed) so the package build, which does not use Vite, resolves them.
import { mergeStringModules } from '../merge';
import * as commands from './commands';
import * as grid from './grid';
import * as shell from './shell';

export const de: Readonly<Record<string, string>> = mergeStringModules({
	'./commands.ts': commands,
	'./grid.ts': grid,
	'./shell.ts': shell,
});
