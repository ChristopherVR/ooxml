<script lang="ts">
	/** One File > Open place for a cloud provider: the shared `pptx-ui-cloud-files` browser. */
	import { mountCloudFilesPlace } from 'ooxml-ui/pptx';
	import type { CloudFileProvider, CloudFilesPlace, CloudOpenedFile } from 'ooxml-ui/pptx';
	import { untrack } from 'svelte';
	import { useTranslator } from '../../../../i18n/context';

	const { provider, onopen }: { provider: CloudFileProvider; onopen: (file: CloudOpenedFile) => void } = $props();
	const t = useTranslator();
	// eslint-disable-next-line no-unassigned-vars
	let hostEl: HTMLDivElement | undefined;

	// Mounted once per provider; the open event always reaches the latest `onopen` (read lazily),
	// and the translator is bound to the viewer's live locale.
	$effect(() => {
		const host = hostEl;
		if (!host) {
			return;
		}
		const place: CloudFilesPlace = mountCloudFilesPlace(host, {
			provider,
			translate: t,
			onOpen: (file) => untrack(() => onopen)(file),
		});
		return () => place.destroy();
	});
</script>

<div bind:this={hostEl} class="place" data-pptx-cloud-place={provider.id}></div>

<style>
	.place {
		max-width: 900px;
		margin-top: 24px;
	}
</style>
