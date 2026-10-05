<script lang="ts">
	import { onMount } from 'svelte';
	import {
		applyTeamsProps,
		defineTeamsApp,
		listenTeamsEvents,
		type TeamsApp,
		type TeamsProps,
	} from 'teams-viewer';

	let props: TeamsProps = $props();
	let el: TeamsApp | undefined = $state();
	let applied: TeamsProps = {};

	onMount(() => {
		defineTeamsApp();
		return listenTeamsEvents(el!, () => props);
	});
	$effect(() => {
		if (!el) return;
		const next: TeamsProps = { ...props };
		applyTeamsProps(el, next, applied);
		applied = next;
	});
</script>

<teams-app bind:this={el} style="display:block;height:100%"></teams-app>
