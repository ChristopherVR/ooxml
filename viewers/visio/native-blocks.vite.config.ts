import { defineConfig } from 'vite';
import { createRequire } from 'node:module';
import { pathToFileURL } from 'node:url';
const require = createRequire(new URL('./packages/bindings/package.json', import.meta.url));
const { svelte } = await import(
	pathToFileURL(require.resolve('@sveltejs/vite-plugin-svelte')).href
);
const { default: solid } = await import(pathToFileURL(require.resolve('vite-plugin-solid')).href);
import { resolve } from 'node:path';
export default defineConfig({
	root: resolve('packages/bindings/tests'),
	plugins: [svelte(), solid({ hot: false })],
	oxc: { decorator: { legacy: true } },
	server: { fs: { allow: [resolve('../..')] } },
});
