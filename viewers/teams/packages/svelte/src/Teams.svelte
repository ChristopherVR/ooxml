<script lang="ts">
	import { onMount } from 'svelte';
	import {
		bindTeams,
		defineTeamsApp,
		pickTeamsProps,
		type TeamsApp,
		type TeamsBinding,
		type TeamsElementProps,
	} from 'teams-viewer';

	// `class` is the host class (`className` works too).
	let props: TeamsElementProps & { class?: string } = $props();
	let el: TeamsApp | undefined = $state();
	let binding: TeamsBinding | undefined;

	onMount(() => {
		defineTeamsApp();
		binding = bindTeams(el!, () => props);
		return () => binding?.destroy();
	});
	$effect(() => {
		binding?.update(pickTeamsProps({ ...props, className: props.class ?? props.className }));
	});
</script>

<teams-app bind:this={el} style="display:block;height:100%"></teams-app>
