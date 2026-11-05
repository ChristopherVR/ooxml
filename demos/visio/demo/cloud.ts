import type { CloudFileProvider } from 'ooxml-core/embedding';
import { mergeDrawingSnapshots } from 'ooxml-core/visio/collab';
import { createDemoOneDriveProvider } from 'ooxml-demo-onedrive';

/**
 * The demos' File > Open places: OneDrive, signed in through the shared demo dialog. A real host
 * passes its own `fileProviders`. Concurrent edits to a drawing are combined by the Visio merger.
 */
export function createDemoFileProviders(): readonly CloudFileProvider[] {
	return [
		createDemoOneDriveProvider({
			mergers: {
				vsdx: ({ base, local, remote }) => mergeDrawingSnapshots(base.bytes, local, remote.bytes),
			},
		}),
	];
}
