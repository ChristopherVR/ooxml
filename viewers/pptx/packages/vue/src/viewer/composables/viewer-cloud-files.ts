/**
 * The cloud places of File > Open (the `fileProviders` prop) and the two things
 * the chrome tells the viewer about them: a file was opened from a place, and
 * another document replaced it. `PowerPointViewer.vue` provides it, so the
 * backstage needs no prop threading; a component rendered standalone sees no
 * providers.
 */
import type { CloudFileProvider, CloudOpenedFile } from 'ooxml-ui/pptx';
import { inject } from 'vue';
import type { InjectionKey } from 'vue';

export interface ViewerCloudFiles {
	/** A getter, so the backstage follows the host's prop. */
	providers: () => readonly CloudFileProvider[];
	open: (file: CloudOpenedFile) => void;
	clear: () => void;
}

const NONE: ViewerCloudFiles = {
	providers: () => [],
	open: () => undefined,
	clear: () => undefined,
};

export const ViewerCloudFilesKey: InjectionKey<ViewerCloudFiles> =
	Symbol('pptx-viewer-cloud-files');

export function useViewerCloudFiles(): ViewerCloudFiles {
	return inject(ViewerCloudFilesKey, NONE);
}
