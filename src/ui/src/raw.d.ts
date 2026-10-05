// Component styles live in real `.css` files next to their element and are imported as text:
// `import css from './chat-list.css?raw'`. Vite and Vitest support the `?raw` suffix natively; the
// package build maps it with a small esbuild plugin (see tsup.config.ts).
declare module '*.css?raw' {
	const css: string;
	export default css;
}
