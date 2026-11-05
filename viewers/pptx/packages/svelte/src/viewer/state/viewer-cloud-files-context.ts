import type { CloudFileProvider, CloudOpenedFile } from 'ooxml-ui/pptx';
import { getContext, setContext } from 'svelte';

/**
 * The cloud places of File > Open (the `fileProviders` prop) and the two things the chrome tells
 * the viewer about them: a file was opened from a place, and another document replaced it.
 * `createViewerState` provides a getter; a component mounted stand-alone sees no providers.
 */
export interface ViewerCloudFiles {
	readonly providers: readonly CloudFileProvider[];
	open(file: CloudOpenedFile): void;
	clear(): void;
}

const VIEWER_CLOUD_FILES_CONTEXT_KEY = Symbol('pptx-svelte-viewer-cloud-files');

const NONE: ViewerCloudFiles = { providers: [], open: () => undefined, clear: () => undefined };

/** Provide the cloud files getter to the component subtree (root component only). */
export function provideViewerCloudFiles(getCloudFiles: () => ViewerCloudFiles): void {
	setContext(VIEWER_CLOUD_FILES_CONTEXT_KEY, getCloudFiles);
}

/** A getter for the viewer's cloud places (no providers outside a viewer). */
export function useViewerCloudFiles(): () => ViewerCloudFiles {
	return (
		getContext<(() => ViewerCloudFiles) | undefined>(VIEWER_CLOUD_FILES_CONTEXT_KEY) ?? (() => NONE)
	);
}
