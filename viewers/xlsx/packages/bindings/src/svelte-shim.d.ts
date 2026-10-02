// Lets `tsc` (and Vitest's type-aware editors) import the Svelte adapter; svelte-check reads the
// real component types from the `.svelte` file itself.
declare module '*.svelte' {
	import type { Component } from 'svelte';
	const component: Component<
		Record<string, unknown>,
		Record<string, (...args: never[]) => unknown>
	>;
	export default component;
}
