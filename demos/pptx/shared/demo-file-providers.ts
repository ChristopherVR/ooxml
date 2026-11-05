/**
 * The cloud places every pptx demo offers under File > Open: OneDrive, signed in through the
 * demo-only sign-in of `ooxml-demo-onedrive`. Concurrent edits to a deck merge with the core's
 * PowerPoint snapshot merge. A real host passes its own provider (docs/onedrive-quickstart.md).
 */
import { mergePptxSnapshots } from 'ooxml-core/pptx';
import { createDemoOneDriveProvider } from 'ooxml-demo-onedrive';

export const demoFileProviders = [
	createDemoOneDriveProvider({
		mergers: {
			pptx: ({ base, local, remote }) => mergePptxSnapshots(base.bytes, local, remote.bytes),
		},
	}),
];
