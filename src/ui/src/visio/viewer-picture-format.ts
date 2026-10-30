import { canEditDiagram } from './controller';
import type { VisioEdit } from 'ooxml-core/visio';
import type { ViewerController, ViewerState } from './controller';
import type { VisioPictureAction } from './ribbon-action';
import type { RibbonCommand } from './ribbon-parts';
import { selectedShape } from './shape-inspector';

/** The picture a single selection names, when it is a picture on the current page. */
export function selectedPicture(
	state: ViewerState,
): { pageId: string; shapeId: string } | undefined {
	const page = state.document?.pages[state.pageIndex];
	const shape = selectedShape(state.document, state.selectedShape, state.pageIndex);
	if (!page || state.selectedShapes.length !== 1 || !shape?.image) return undefined;
	return { pageId: page.id, shapeId: shape.id };
}

/**
 * Picture Format > Adjust (Brightness, Contrast) and Crop: one core edit per command, on the
 * selected picture. Core admission decides whether the picture can take it.
 */
export class ViewerPictureFormat {
	constructor(
		private readonly root: ShadowRoot,
		private readonly controller: ViewerController,
		private readonly edit: (run: () => Promise<void>, message: string) => void,
	) {}
	reason(state: ViewerState): string | undefined {
		if (!canEditDiagram(state)) return 'Open a .vsdx file to edit a picture.';
		if (state.loading || state.edit.busy) return 'Wait for the current operation to finish.';
		if (!selectedPicture(state)) return 'Select one picture to adjust or crop it.';
		return undefined;
	}
	run(action: VisioPictureAction): void {
		const state = this.controller.state;
		const target = selectedPicture(state);
		if (!target || this.reason(state) !== undefined) return;
		const edit: VisioEdit =
			action.type === 'picture-crop'
				? { type: 'crop-picture', ...target, crop: action.crop }
				: {
						type: 'set-picture-adjust',
						...target,
						...(action.brightness === undefined ? {} : { brightness: action.brightness }),
						...(action.contrast === undefined ? {} : { contrast: action.contrast }),
					};
		const message =
			action.type === 'picture-crop'
				? action.crop === 'square'
					? 'Cropped the picture to a square.'
					: 'Reset the picture crop.'
				: 'Adjusted the picture.';
		this.edit(() => this.controller.applyEdits([edit]), message);
	}
	render(state: ViewerState): void {
		const reason = this.reason(state);
		for (const id of ['picture-brightness', 'picture-contrast', 'picture-crop']) {
			const menu = this.root.querySelector<RibbonCommand>(`[data-menu="${id}"]`);
			if (!menu) continue;
			menu.disabled = reason !== undefined;
			menu.toggleAttribute('disabled', reason !== undefined);
			menu.title = reason
				? `${menu.getAttribute('label') ?? ''}: ${reason}`
				: (menu.getAttribute('label') ?? '');
		}
	}
}
