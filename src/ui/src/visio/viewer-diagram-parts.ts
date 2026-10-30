import { canEditDiagram } from './controller';
import {
	visioCalloutCommand,
	visioContainerCommand,
	visioSelectionIsOnPage,
} from 'ooxml-core/visio/ui';
import type { OfficeUiGallery } from '../ribbon/gallery';
import type { ViewerController, ViewerState } from './controller';
import type { VisioContainerOptions, VisioEdit } from 'ooxml-core/visio';
import type {
	VisioContainerAction,
	VisioContainerOptionAction,
	VisioDiagramPartAction,
} from './ribbon-action';
import { CONTAINER_MARGINS } from './ribbon-contextual-formats';
import { selectedShape } from './shape-inspector';
import {
	CALLOUT_STYLE_NAMES,
	CONTAINER_STYLE_NAMES,
	syncDiagramPart,
	type DiagramPart,
} from './ribbon-diagram-parts';
import type { RibbonCommand } from './ribbon-parts';

/** The one selected container: its page, shape, members and stored options. */
type ContainerTarget = {
	pageId: string;
	id: string;
	members: string[];
	options: VisioContainerOptions;
};
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
	/** The selected container's page, members and stored options, when exactly one container is selected. */
	#container(
		state: ViewerState,
	): { pageId: string; id: string; members: string[]; options: VisioContainerOptions } | undefined {
		const page = state.document?.pages[state.pageIndex];
		const shape = selectedShape(state.document, state.selectedShape, state.pageIndex);
		if (!page || state.selectedShapes.length !== 1 || shape?.structure?.type !== 'container')
			return undefined;
		return {
			pageId: page.id,
			id: shape.id,
			members: shape.structure.memberIds,
			options: shape.structure.options ?? {},
		};
	}
	/**
	 * Container Format > Margins, Automatic Resize and Lock Container: one core option edit. The
	 * lock toggle reads the container's stored lock, so it flips whatever the page shows.
	 */
	containerOptions(action: VisioContainerOptionAction): void {
		const state = this.controller.state;
		const target = this.#container(state);
		if (!target || !canEditDiagram(state) || state.loading || state.edit.busy) return;
		const locked = action.locked === 'toggle' ? !target.options.locked : action.locked;
		const edit: VisioEdit = {
			type: 'set-container-options',
			pageId: target.pageId,
			shapeId: target.id,
			...(action.margin === undefined ? {} : { margin: action.margin }),
			...(action.resize === undefined ? {} : { resize: action.resize }),
			...(locked === undefined ? {} : { locked }),
		};
		const message =
			locked !== undefined
				? locked
					? 'Locked the container.'
					: 'Unlocked the container.'
				: action.resize !== undefined
					? 'Changed the automatic resize of the container.'
					: 'Changed the container margin.';
		this.edit(() => this.controller.applyEdits([edit]), message);
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
		this.#renderOptions(state, target, editable);
		const item = this.root.querySelector<RibbonCommand>('[command="ctx-container"]');
		if (item) {
			const reason = this.reason(state, 'container');
			item.disabled = reason !== undefined;
			item.title = reason ? `Container: ${reason}` : 'Container: frame the selected shapes.';
		}
	}
	/** Margins, Automatic Resize and Lock Container: enabled with a container, checked from its options. */
	#renderOptions(state: ViewerState, target: ContainerTarget | undefined, editable: boolean): void {
		const tab = this.root.querySelector('[data-ribbon-tab="container-format"]');
		const reason = target ? undefined : 'Select a container to change its options.';
		const enabled = !!target && editable && !state.loading;
		for (const [id, label] of [
			['container-margins', 'Margins'],
			['container-resize', 'Automatic Resize'],
		] as const) {
			const menu = tab?.querySelector<RibbonCommand>(`[data-menu="${id}"]`);
			if (!menu) continue;
			menu.disabled = !enabled;
			menu.toggleAttribute('disabled', !enabled);
			menu.title = enabled
				? label
				: `${label}: ${reason ?? 'Open a .vsdx file to change the container.'}`;
		}
		const lock = tab?.querySelector<RibbonCommand>('[command="container-lock"]');
		if (lock) {
			lock.disabled = !enabled;
			lock.toggleAttribute('disabled', !enabled);
			lock.setAttribute('checked', String(!!target?.options.locked));
		}
		for (const item of tab?.querySelectorAll<HTMLElement>('[command^="container-resize-"]') ?? []) {
			const mode = Number(item.getAttribute('command')!.slice('container-resize-'.length));
			item.setAttribute('checked', String(mode === (target?.options.resize ?? 1)));
		}
		CONTAINER_MARGINS.forEach(([, margin], index) => {
			const item = tab?.querySelector<HTMLElement>(`[command="container-margin-${index}"]`);
			item?.setAttribute(
				'checked',
				String(Math.abs((target?.options.margin ?? 0.25) - margin) < 1e-6),
			);
		});
	}
}
