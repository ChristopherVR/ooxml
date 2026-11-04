import { createPptxViewer } from 'pptx-vanilla-viewer';
import type { PptxViewerInstance } from 'pptx-vanilla-viewer';

import { mountOneDriveDemo } from '../../shared/onedrive-demo';

/** `?onedrive=1`: open, edit and save a deck on your own OneDrive (docs/onedrive-quickstart.md). */
export function mountOneDriveDemoApp(root: HTMLElement): void {
	const bar = document.createElement('div');
	const shell = document.createElement('main');
	shell.className = 'od-viewer';
	const empty = document.createElement('p');
	empty.className = 'od-empty';
	empty.textContent = 'No OneDrive file open yet.';
	shell.append(empty);
	root.replaceChildren(bar, shell);
	let viewer: PptxViewerInstance | undefined;
	// A new viewer per deck: a newly opened file, a merge, or "open latest".
	const mounted = mountOneDriveDemo(bar, (deck) => {
		viewer?.destroy();
		shell.replaceChildren();
		viewer = createPptxViewer(shell, {
			source: deck.bytes,
			fileName: deck.fileName,
			editable: deck.canEdit,
			storage: mounted.storage,
		});
	});
	import.meta.hot?.dispose(() => {
		viewer?.destroy();
		mounted.dispose();
	});
}
