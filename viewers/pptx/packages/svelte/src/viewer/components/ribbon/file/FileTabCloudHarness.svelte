<script lang="ts">
	/**
	 * Test-only shell for `FileTab`: provides the viewer's cloud files context (as
	 * `createViewerState` does) when `cloud` is given, then renders the File tab.
	 */
	import { provideViewerCloudFiles } from '../../../state/viewer-cloud-files-context';
	import type { ViewerCloudFiles } from '../../../state/viewer-cloud-files-context';
	import FileTab from './FileTab.svelte';

	const {
		cloud,
		onclose,
		oncreatepresentation,
	}: {
		cloud?: ViewerCloudFiles | undefined;
		onclose: () => void;
		oncreatepresentation: (templateId: string) => void;
	} = $props();
	// svelte-ignore state_referenced_locally
	const provided = cloud;
	if (provided) {
		provideViewerCloudFiles(() => provided);
	}
	const noop = (): void => undefined;
</script>

<FileTab {onclose} {oncreatepresentation} ondownload={noop} ondownloadppsx={noop} ondownloadpptm={noop} ondownloadppt={noop} hasMacros={false} />
