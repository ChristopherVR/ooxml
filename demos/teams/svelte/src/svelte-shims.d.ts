// tsc cannot read .svelte files; the demo only needs the component to be mountable.
declare module 'openteams-svelte-viewer' {
	import type { Component } from 'svelte';
	const Teams: Component<Record<string, unknown>>;
	export default Teams;
}
