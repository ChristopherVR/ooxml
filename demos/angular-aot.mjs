import { createRequire, registerHooks } from 'node:module';
import { pathToFileURL } from 'node:url';

/**
 * The Angular compiler plugin for an Angular demo, compiling ahead of time (`jit: false`): the
 * published site's Content Security Policy never allows `eval`, which Angular's run-time (JIT)
 * compiler needs. The demo's tsconfig must reach the binding's source, so its decorators are
 * compiled too.
 *
 * `packageJson` is the demo package that declares `@analogjs/vite-plugin-angular` and TypeScript 6.
 * The plugin loads `typescript` without declaring it, so under the workspace's isolated linker it
 * would resolve Bun's shared fallback: the core's TypeScript 7, the native port without the JS
 * compiler API the plugin calls. Its imports are pointed at that package's TypeScript, then the
 * plugin is loaded (dynamically, so the hook is in place first).
 *
 * @param {string} packageJson
 * @param {string} tsconfig
 * @returns {Promise<import('vite').Plugin[]>}
 */
export async function angularAot(packageJson, tsconfig) {
	const require = createRequire(packageJson);
	const typescript = pathToFileURL(require.resolve('typescript')).href;
	registerHooks({
		resolve(specifier, context, next) {
			if (specifier === 'typescript' && context.parentURL?.includes('/@analogjs/')) {
				return { url: typescript, format: 'commonjs', shortCircuit: true };
			}
			return next(specifier, context);
		},
	});
	const plugin = pathToFileURL(require.resolve('@analogjs/vite-plugin-angular')).href;
	const { default: angular } = await import(plugin);
	return angular({ tsconfig, jit: false });
}
