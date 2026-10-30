import { canEditDiagram } from './controller';
import {
	visioCalloutCommand,
	visioContainerCommand,
	visioSelectionIsOnPage,
} from 'ooxml-core/visio/ui';
import type { OfficeUiGallery } from '../ribbon/gallery';
import type { ViewerController, ViewerState } from './controller';
import type { VisioContainerAction, VisioDiagramPartAction } from './ribbon-action';
import { selectedShape } from './shape-inspector';
import {
	CALLOUT_STYLE_NAMES,
	CONTAINER_STYLE_NAMES,
	syncDiagramPart,
	type DiagramPart,
} from './ribbon-diagram-parts';
import type { RibbonCommand } from './ribbon-parts';

/**
 * Insert > Container and Callout (and the shape menu's Container): one core edit each, so undo
 * restores the page in one step. Availability follows the selection; core admission decides.
 */
export class ViewerDiagramParts {
	constructor(
		private readonly root: ShadowRoot,
		private readonly controller: ViewerController,
		private readonly edit: (run: () => Promise<void>, message: string) => void,
		private readonly revealText: () => void,
	) {}
	reason(state: ViewerState, part: DiagramPart): string | undefined {
		if (!canEditDiagram(state)) return 'Open a .vsdx file to insert diagram parts.';
		if (state.loading || state.edit.busy) return 'Wait for the current operation to finish.';
		const page = state.document?.pages[state.pageIndex];
		if (!page || state.selectedShapes.some((shape) => !visioSelectionIsOnPage(shape, page.id)))
			return 'Select shapes on the current page.';
		const ids = state.selectedShapes.map((shape) => shape.id);
		if (part === 'container' && !visioContainerCommand(page, ids, 'classic'))
			return 'Select top-level shapes on this page, or nothing for an empty container.';
		if (part === 'callout' && !visioCalloutCommand(page, ids, 'rectangle'))
			return 'Select one two-dimensional shape to attach the callout to.';
		return undefined;
	}
	run(action: VisioDiagramPartAction): void {
		if (this.reason(this.controller.state, action.part) !== undefined) return;
		const name =
			action.part === 'container'
				? `${CONTAINER_STYLE_NAMES[action.style]} container`
				: `${CALLOUT_STYLE_NAMES[action.style]} callout`;
		this.edit(async () => {
			await this.controller.insertDiagramPart(
				action.part === 'container'
					? { kind: 'container', style: action.style }
					: { kind: 'callout', style: action.style },
			);
			// A new callout is empty: type its text straight away, as in Visio.
			if (action.part === 'callout') this.revealText();
		}, `Inserted a ${name}.`);
	}
	/** The selected container's page and members, when exactly one container is selected. */
	#container(state: ViewerState): { pageId: string; id: string; members: string[] } | undefined {
		const page = state.document?.pages[state.pageIndex];
		const shape = selectedShape(state.document, state.selectedShape, state.pageIndex);
		if (!page || state.selectedShapes.length !== 1 || shape?.structure?.type !== 'container')
			return undefined;
		return { pageId: page.id, id: shape.id, members: shape.structure.memberIds };
	}
	/** Container Format > Fit to Contents, Select Contents and Disband Container. */
	container(command: VisioContainerAction['command']): void {
		const state = this.controller.state;
		const target = this.#container(state);
		if (!target || (command !== 'select-contents' && !canEditDiagram(state))) return;
		if (command === 'select-contents') {
			this.controller.selectShapes(
				target.members.map((id) => ({ id, name: id, pageId: target.pageId })),
			);
			return;
		}
		if (state.loading || state.edit.busy) return;
		if (command === 'fit')
			this.edit(
				() =>
					this.controller.applyEdits([
						{ type: 'reconcile-containers', pageId: target.pageId, shapeIds: target.members },
					]),
				'Fitted the container to its contents.',
			);
		else
			// Visio's Disband removes the frame and leaves its members on the page.
			this.edit(
				() =>
					this.controller.applyEdits([
						{ type: 'delete-shape', pageId: target.pageId, shapeId: target.id },
					]),
				'Disbanded the container; its contents stay on the page.',
			);
	}
	render(state: ViewerState): void {
		const target = this.#container(state);
		const editable = canEditDiagram(state) && !state.loading && !state.edit.busy;
		const tab = this.root.querySelector('[data-ribbon-tab="container-format"]');
		for (const [id, enabled, reason] of [
			[
				'container-fit',
				!!target?.members.length && editable,
				'Open a .vsdx file with members to fit.',
			],
			['container-disband', !!target && editable, 'Open a .vsdx file to disband a container.'],
			['container-select-contents', !!target?.members.length, 'This container has no members.'],
		] as const) {
			const control = tab?.querySelector<RibbonCommand>(`[command="${id}"]`);
			if (!control) continue;
			control.disabled = !enabled;
			control.toggleAttribute('disabled', !enabled);
			const label = control.getAttribute('label') ?? '';
			control.title = enabled ? label : `${label}: ${reason}`;
		}
		for (const part of ['container', 'callout'] as const) {
			const gallery = this.root.querySelector<OfficeUiGallery>(
				`office-ui-gallery[command="${part}"]`,
			);
			if (gallery) syncDiagramPart(gallery, part, this.reason(state, part));
		}
		const item = this.root.querySelector<RibbonCommand>('[command="ctx-container"]');
		if (item) {
			const reason = this.reason(state, 'container');
			item.disabled = reason !== undefined;
			item.title = reason ? `Container: ${reason}` : 'Container: frame the selected shapes.';
		}
	}
}
