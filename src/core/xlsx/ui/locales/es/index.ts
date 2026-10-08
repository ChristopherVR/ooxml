// Every string file of this folder merged into one table; see ../merge.ts. Listed explicitly
// (not globbed) so the package build, which does not use Vite, resolves them.
import { mergeStringModules } from '../merge';
import * as collab from './collab';
import * as commands from './commands';
import * as grid from './grid';
import * as shell from './shell';

export const es: Readonly<Record<string, string>> = mergeStringModules({
	'./collab.ts': collab,
	'./commands.ts': commands,
	'./grid.ts': grid,
	'./shell.ts': shell,
});
