import { fileURLToPath } from 'node:url';
import { defineConfig } from 'vite';
import { angularAot } from '../../angular-aot.mjs';
import { sharedConfig } from '../vite.shared.ts';

const here = (path: string): string => fileURLToPath(new URL(path, import.meta.url));

// Compiled ahead of time: tsconfig.json maps the binding to its source for the Angular compiler.
export default defineConfig({
	...sharedConfig(5178),
	plugins: [await angularAot(here('package.json'), here('tsconfig.json'))],
	// The Angular plugin turns Vite's TypeScript transform off unless the config sets one; the dev
	// server still needs it for the core and UI sources, which are outside the Angular program.
	oxc: {},
});
