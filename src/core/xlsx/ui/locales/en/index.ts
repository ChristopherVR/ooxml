// Every string file of this folder merged into one table; see ../merge.ts. Listed explicitly
// (not globbed) so the package build, which does not use Vite, resolves them.
import { mergeStringModules } from '../merge';
import * as collab from './collab';
import * as commands from './commands';
import * as grid from './grid';
import * as options from './options';
import * as shell from './shell';
import * as smartart from './smartart';

export const en: Readonly<Record<string, string>> = mergeStringModules({
	'./collab.ts': collab,
	'./commands.ts': commands,
	'./grid.ts': grid,
	'./options.ts': options,
	'./shell.ts': shell,
	'./smartart.ts': smartart,
});
