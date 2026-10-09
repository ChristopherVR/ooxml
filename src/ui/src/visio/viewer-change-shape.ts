import type { VisioChangeShapeTarget } from 'ooxml-core/visio';
import { visioChangeShapeCommand, visioChangeShapeRefusal } from 'ooxml-core/visio/ui';
import type { OfficeUiGallery } from '../ribbon/gallery';
import type { ViewerController, ViewerState } from './controller';
import { syncChangeShape } from './ribbon-change-shape';
import { BASIC_SHAPES } from './shapes-window';

/**
 * Home > Change Shape: replaces the outline of the one selected local shape through the core's
 * `change-shape` edit, so undo and redo come from the source-backed history. The gallery is
 * enabled only for a selection the scene admits; locks and formulas are proven by the core.
 */
export class ViewerChangeShape {
	constructor(
		private readonly root: ShadowRoot,
		private readonly controller: ViewerController,
		private readonly edit: (run: () => Promise<void>, message: string) => void,
	) {}
	#refusal(state: ViewerState): string | undefined {
		if (!state.edit.sourceAvailable) return 'Open a .vsdx file to change shapes.';
		if (state.loading || state.edit.busy) return 'Wait for the current edit to finish.';
		const page = state.document?.pages[state.pageIndex];
		if (state.selectedShapes.some((shape) => shape.pageId && shape.pageId !== page?.id))
			return 'Select a shape on this page.';
		return visioChangeShapeRefusal(page, state.selectedShapes);
	}
	run(shape: VisioChangeShapeTarget): void {
		const state = this.controller.state;
		if (this.#refusal(state) !== undefined) return;
		const command = visioChangeShapeCommand(
			state.document?.pages[state.pageIndex],
			state.selectedShapes,
			shape,
		);
		const name = BASIC_SHAPES.find((master) => master.id === shape)?.name ?? shape;
		if (command)
			this.edit(() => this.controller.applyEdits([command]), `Changed the shape to ${name}.`);
	}
	render(state: ViewerState): void {
		const gallery = this.root.querySelector<OfficeUiGallery>(
			'office-ui-gallery[command="change-shape"]',
		);
		if (gallery) syncChangeShape(gallery, this.#refusal(state));
	}
}
