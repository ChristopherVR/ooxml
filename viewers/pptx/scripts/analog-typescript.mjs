/**
 * analog-typescript.mjs: give `@analogjs/vite-plugin-angular` the Angular demo's TypeScript.
 *
 * The plugin loads `typescript` without declaring it, so under the root workspace's isolated
 * linker it resolves Bun's shared fallback, which holds the core's TypeScript 7 (the native
 * port, without the JS compiler API the plugin calls). Loaded with
 * `node --import <this file>` from the demo's folder, it points the plugin's `typescript`
 * imports and requires at the version the demo itself depends on.
 */

import { createRequire, registerHooks } from 'node:module';
import { join } from 'node:path';
import { pathToFileURL } from 'node:url';

const typescript = pathToFileURL(
	createRequire(join(process.cwd(), 'package.json')).resolve('typescript'),
).href;

registerHooks({
	resolve(specifier, context, next) {
		if (specifier === 'typescript' && context.parentURL?.includes('/@analogjs/')) {
			return { url: typescript, format: 'commonjs', shortCircuit: true };
		}
		return next(specifier, context);
	},
});
