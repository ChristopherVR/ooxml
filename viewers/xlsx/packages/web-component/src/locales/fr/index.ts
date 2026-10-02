// Every string file of this folder (shell.ts from the shell, grid.ts from the grid, commands.ts
// from the commands) merged into one table; see ../merge.ts.
import { mergeStringModules } from '../merge';

const modules = import.meta.glob<Record<string, unknown>>(
	['./*.ts', '!./index.ts', '!./*.test.ts'],
	{
		eager: true,
	},
);

export const fr: Readonly<Record<string, string>> = mergeStringModules(modules);
