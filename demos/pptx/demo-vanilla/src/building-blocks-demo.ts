import { createViewerBuildingBlocks } from 'pptx-vanilla-viewer';
import type { ViewerBuildingBlocksOptions } from 'pptx-vanilla-viewer';

import { mountBuildingBlocksPanel } from '../../shared/building-blocks-panel';

/**
 * `?blocks=1`: a custom shell composed from the vanilla building blocks (toolbar, slides pane,
 * canvas) with the shared host panel that finishes their workflows. The framework-neutral
 * `e2e/pptx/building-blocks.spec.ts` drives it.
 */
export async function mountBuildingBlocksDemo(root: HTMLElement): Promise<void> {
	const response = await fetch(`${import.meta.env.BASE_URL}sample-deck.pptx`);
	const content = new Uint8Array(await response.arrayBuffer());
	let options: ViewerBuildingBlocksOptions = {
		content,
		canEdit: false,
		fileName: 'sample-deck.pptx',
	};
	const blocks = createViewerBuildingBlocks(options);
	const panel = document.createElement('section');
	const toolbar = document.createElement('div');
	const body = document.createElement('div');
	body.style.cssText = 'display:flex;height:calc(100vh - 160px)';
	const rail = document.createElement('div');
	// The rail's styles are scoped under a `.pptxv` ancestor.
	rail.className = 'pptxv';
	rail.style.cssText = 'width:200px;overflow:auto';
	const canvas = document.createElement('div');
	canvas.style.cssText = 'flex:1;position:relative';
	body.append(rail, canvas);
	root.replaceChildren(panel, toolbar, body);
	const host = mountBuildingBlocksPanel(panel, {
		actions: () => blocks.actions,
		state: () => blocks.getState(),
		setCanEdit: (canEdit) => {
			options = { ...options, canEdit };
			blocks.update(options);
		},
		setCollaboration: (collaboration) => {
			options = {
				...options,
				collaboration: collaboration as ViewerBuildingBlocksOptions['collaboration'],
			};
			blocks.update(options);
		},
	});
	blocks.subscribe(() => host.refresh());
	blocks.mountToolbar(toolbar);
	blocks.mountSlidesPane(rail);
	blocks.mountCanvas(canvas);
}
