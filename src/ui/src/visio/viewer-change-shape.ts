import type { VisioChangeShapeTarget, VisioMaster } from 'ooxml-core/visio';
import {
	visioChangeMasterCommand,
	visioChangeMasterTargets,
	visioChangeShapeCommand,
	visioChangeShapeRefusal,
} from 'ooxml-core/visio/ui';
import type { OfficeUiGallery } from '../ribbon/gallery';
import type { ViewerController, ViewerState } from './controller';
import { syncChangeShape, type ChangeShapeMaster } from './ribbon-change-shape';
import { DOCUMENT_MASTER_PREFIX, shapesDocument } from './shapes-document';
import { BASIC_SHAPES } from './shapes-window';

/** What the gallery asks for: a Basic Shapes outline, or `document:<id>` for a master. */
export type ChangeShapeChoice =
	| VisioChangeShapeTarget
	| `${typeof DOCUMENT_MASTER_PREFIX}${string}`;

/**
 * Home > Change Shape through the core's `change-shape` edit, so undo and redo come from the
 * source-backed history. A shape drawn here takes another Basic Shapes outline; a stencil shape
 * becomes an instance of another master of the Document Stencil, as Visio's Change Shape does.
 * The gallery shows the choices for the selection; locks and formulas are proven by the core.
 */
export class ViewerChangeShape {
	#listed = '';
	#masters: readonly ChangeShapeMaster[] = [];
	#shown:
		| { reason: string | undefined; tiles: readonly ChangeShapeMaster[] | undefined }
		| undefined;
	constructor(
		private readonly root: ShadowRoot,
		private readonly controller: ViewerController,
		private readonly edit: (run: () => Promise<void>, message: string) => void,
	) {}
	#busy(state: ViewerState): string | undefined {
		if (!state.edit.sourceAvailable) return 'Open a .vsdx file to change shapes.';
		if (state.loading || state.edit.busy) return 'Wait for the current edit to finish.';
		const page = state.document?.pages[state.pageIndex];
		if (state.selectedShapes.some((shape) => shape.pageId && shape.pageId !== page?.id))
			return 'Select a shape on this page.';
		return undefined;
	}
	/** The masters the selected stencil shape can change to, or the reason it cannot. */
	#targets(state: ViewerState) {
		return visioChangeMasterTargets(
			state.document,
			state.document?.pages[state.pageIndex],
			state.selectedShapes,
		);
	}
	run(shape: ChangeShapeChoice): void {
		const state = this.controller.state;
		if (this.#busy(state) !== undefined) return;
		const page = state.document?.pages[state.pageIndex];
		if (shape.startsWith(DOCUMENT_MASTER_PREFIX)) {
			const masterId = shape.slice(DOCUMENT_MASTER_PREFIX.length);
			const command = visioChangeMasterCommand(
				state.document,
				page,
				state.selectedShapes,
				masterId,
			);
			const name = state.document?.masters?.find((master) => master.id === masterId)?.name;
			if (command)
				this.edit(
					() => this.controller.applyEdits([command]),
					`Changed the shape to ${name ?? 'the master'}.`,
				);
			return;
		}
		if (visioChangeShapeRefusal(page, state.selectedShapes) !== undefined) return;
		const command = visioChangeShapeCommand(
			page,
			state.selectedShapes,
			shape as VisioChangeShapeTarget,
		);
		const name = BASIC_SHAPES.find((master) => master.id === shape)?.name ?? shape;
		if (command)
			this.edit(() => this.controller.applyEdits([command]), `Changed the shape to ${name}.`);
	}
	/** The tiles for `targets`, drawn once per list of masters: an edit reparses the drawing. */
	#tiles(state: ViewerState, targets: readonly VisioMaster[]): readonly ChangeShapeMaster[] {
		const key = JSON.stringify(targets.map((master) => [master.id, master.name]));
		if (key !== this.#listed) {
			const drawn = new Map(
				shapesDocument(state.document).masters.map((master) => [master.id, master.draw]),
			);
			this.#listed = key;
			this.#masters = targets.map((master) => {
				const svg = drawn.get(`${DOCUMENT_MASTER_PREFIX}${master.id}`)?.();
				return {
					id: master.id,
					name: master.name,
					...(svg ? { preview: new XMLSerializer().serializeToString(svg) } : {}),
				};
			});
		}
		return this.#masters;
	}
	render(state: ViewerState): void {
		const gallery = this.root.querySelector<OfficeUiGallery>(
			'office-ui-gallery[command="change-shape"]',
		);
		if (!gallery) return;
		const busy = this.#busy(state);
		const page = state.document?.pages[state.pageIndex];
		const stencil =
			state.selectedShapes.length === 1 &&
			!!page?.shapes.find((shape) => shape.id === state.selectedShapes[0]!.id)?.masterId;
		if (busy !== undefined || !stencil) {
			this.#shown = undefined;
			return syncChangeShape(gallery, busy ?? visioChangeShapeRefusal(page, state.selectedShapes));
		}
		const targets = this.#targets(state);
		const tiles = targets.masters.length ? this.#tiles(state, targets.masters) : undefined;
		// The tiles carry drawn previews: hand them to the gallery only when they change.
		const shown = this.#shown;
		if (shown && shown.reason === targets.reason && shown.tiles === tiles) return;
		this.#shown = { reason: targets.reason, tiles };
		syncChangeShape(gallery, targets.reason, tiles);
	}
}
