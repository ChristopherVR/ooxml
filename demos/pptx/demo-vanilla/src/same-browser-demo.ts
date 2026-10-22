import { createPptxViewer } from 'pptx-vanilla-viewer';
import type { PptxViewerInstance } from 'pptx-vanilla-viewer';

import { createSameBrowserDemo } from '../../shared/same-browser-collaboration';

export async function mountSameBrowserDemo(root: HTMLElement): Promise<void> {
	const demo = await createSameBrowserDemo();
	const shell = document.createElement('main');
	shell.style.cssText = 'position:fixed;inset:28px 0 0';
	root.replaceChildren(shell);
	const viewer: PptxViewerInstance = createPptxViewer(shell, {
		source: demo.source,
		fileName: demo.fileName,
		collaboration: demo.collaboration,
		editable: true,
		onCollaborationChange: demo.report,
	});
	import.meta.hot?.dispose(() => {
		viewer.destroy();
		demo.dispose();
	});
}
