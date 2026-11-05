import type { CloudFileProvider, CloudOpenedFile } from 'ooxml-ui/pptx';
import { createContext, useContext } from 'react';

/**
 * The cloud places of File > Open (the `fileProviders` prop) and the two things the chrome tells
 * the viewer about them: a file was opened from a place, and another document replaced it.
 * `PowerPointViewer` provides it; standalone consumers see no providers.
 */
export interface ViewerCloudFiles {
	providers: readonly CloudFileProvider[];
	open: (file: CloudOpenedFile) => void;
	clear: () => void;
}

const NONE: ViewerCloudFiles = { providers: [], open: () => undefined, clear: () => undefined };

export const ViewerCloudFilesContext = createContext<ViewerCloudFiles>(NONE);

export function useViewerCloudFiles(): ViewerCloudFiles {
	return useContext(ViewerCloudFilesContext);
}
