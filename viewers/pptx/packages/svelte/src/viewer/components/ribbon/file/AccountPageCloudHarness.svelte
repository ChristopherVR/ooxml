<script lang="ts">
	/**
	 * Test-only shell for `AccountPage`: provides the viewer's cloud files context (as
	 * `createViewerState` does) when `cloud` is given, then renders the Account page.
	 */
	import { provideViewerCloudFiles } from '../../../state/viewer-cloud-files-context';
	import type { ViewerCloudFiles } from '../../../state/viewer-cloud-files-context';
	import AccountPage from './AccountPage.svelte';

	const { cloud }: { cloud?: ViewerCloudFiles | undefined } = $props();
	// svelte-ignore state_referenced_locally
	const provided = cloud;
	if (provided) {
		provideViewerCloudFiles(() => provided);
	}
</script>

<AccountPage />
