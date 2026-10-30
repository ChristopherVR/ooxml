import type { VisioEdit } from 'ooxml-core/visio';
import {
	visioDeleteSwimlane,
	visioInsertSeparator,
	visioInsertSwimlane,
	visioResizeSwimlane,
	visioSelectionIsOnPage,
	visioSwimlaneInfo,
	visioSwimlaneOptions,
	type VisioSwimlaneInfo,
} from 'ooxml-core/visio/ui';
import type { VisioPage } from 'ooxml-core/visio';
import { canEditDiagram, type ViewerController, type ViewerState } from './controller';
import type { VisioSwimlaneAction } from './ribbon-action';
import type { RibbonCommand } from './ribbon-parts';

type Command = VisioSwimlaneAction['command'];

/** The flowchart the first selected shape belongs to, and the lane it is (or sits in). */
export function swimlaneSelection(
	state: ViewerState,
): { info: VisioSwimlaneInfo; pageId: string; lane?: string; isLane: boolean } | undefined {
	const page = state.document?.pages[state.pageIndex];
	const first = state.selectedShapes[0];
	if (!page || !first || !visioSelectionIsOnPage(first, page.id)) return undefined;
	const info = visioSwimlaneInfo(page, first.id);
	if (!info) return undefined;
	const own = info.lanes.find((lane) => lane.id === first.id);
	const holder = info.lanes.find(
		(lane) => lane.structure?.type === 'container' && lane.structure.memberIds.includes(first.id),
	);
	const lane = own ?? holder;
	return { info, pageId: page.id, isLane: own !== undefined, ...(lane ? { lane: lane.id } : {}) };
}

/**
 * A resize drag on a flowchart's lane or frame. null: the selection is not part of a flowchart,
 * so the ordinary resize applies. undefined: refused (the frame follows its lanes). Otherwise
 * the lane edit that stacks the other lanes and carries their shapes.
 */
export function swimlaneResizeEdits(
	page: VisioPage,
	shapeIds: readonly string[],
	frame: { width: number; height: number },
): VisioEdit[] | null | undefined {
	// Only the frame and the lanes are special; shapes inside a lane resize as usual.
	const parts = shapeIds
		.map((id) => visioSwimlaneInfo(page, id))
		.map((info, index) =>
			info &&
			(info.frame.id === shapeIds[index] || info.lanes.some((l) => l.id === shapeIds[index]))
				? info
				: undefined,
		);
	if (!parts.some(Boolean)) return null;
	const info = parts[0];
	const lane =
		shapeIds.length === 1 ? info?.lanes.find((item) => item.id === shapeIds[0]) : undefined;
	if (!info || !lane) return undefined;
	const size = info.options.direction === 'horizontal' ? frame.height : frame.width;
	return visioResizeSwimlane(page, lane.id, size);
}

/** The tab's controllers: one or two core edits per command, so undo steps back one command. */
export class ViewerSwimlanes {
	constructor(
		private readonly root: ShadowRoot,
		private readonly controller: ViewerController,
		private readonly edit: (run: () => Promise<void>, message: string) => void,
	) {}
	reason(state: ViewerState, command: Command): string | undefined {
		if (!canEditDiagram(state)) return 'Open a .vsdx file to edit the flowchart.';
		if (state.loading || state.edit.busy) return 'Wait for the current operation to finish.';
		const selection = swimlaneSelection(state);
		if (!selection) return 'Select a cross-functional flowchart.';
		if (command === 'delete') {
			if (!selection.isLane) return 'Select the swimlane to delete.';
			if (selection.info.lanes.length < 2) return 'A flowchart keeps at least one swimlane.';
		}
		return undefined;
	}
	#plan(action: VisioSwimlaneAction): { batches: VisioEdit[][]; message: string } | undefined {
		const state = this.controller.state;
		const page = state.document?.pages[state.pageIndex];
		const selection = swimlaneSelection(state);
		if (!page || !selection) return undefined;
		const { info, lane } = selection;
		const next = `Function ${info.lanes.length + 1}`;
		const one = (edits: VisioEdit[] | undefined, message: string) =>
			edits ? { batches: [edits], message } : undefined;
		switch (action.command) {
			case 'insert-before': {
				const target = lane ?? info.lanes[0]?.id;
				return target
					? one(visioInsertSwimlane(page, target, 'before', next), 'Inserted a swimlane.')
					: undefined;
			}
			case 'insert-after': {
				const target = lane ?? info.lanes[info.lanes.length - 1]?.id;
				return target
					? one(visioInsertSwimlane(page, target, 'after', next), 'Inserted a swimlane.')
					: undefined;
			}
			case 'delete': {
				const batches = lane ? visioDeleteSwimlane(page, lane) : undefined;
				return batches ? { batches, message: 'Deleted the swimlane and its shapes.' } : undefined;
			}
			case 'separator':
				return one(visioInsertSeparator(page, info.frame.id), 'Inserted a separator.');
			case 'title':
				return one(
					visioSwimlaneOptions(page, info.frame.id, { showTitle: !info.options.showTitle }),
					info.options.showTitle ? 'Hid the title bar.' : 'Showed the title bar.',
				);
			case 'separators':
				return one(
					visioSwimlaneOptions(page, info.frame.id, {
						showSeparators: !info.options.showSeparators,
					}),
					info.options.showSeparators ? 'Hid the separator bars.' : 'Showed the separator bars.',
				);
			case 'rotate-labels':
				return one(
					visioSwimlaneOptions(page, info.frame.id, { rotateLabels: !info.options.rotateLabels }),
					'Rotated the lane labels.',
				);
			case 'style':
				return action.style === undefined
					? undefined
					: one(
							visioSwimlaneOptions(page, info.frame.id, { style: action.style }),
							`Applied style ${action.style}.`,
						);
		}
	}
	run(action: VisioSwimlaneAction): void {
		if (this.reason(this.controller.state, action.command) !== undefined) return;
		const plan = this.#plan(action);
		if (!plan) return;
		this.edit(async () => {
			for (const batch of plan.batches) await this.controller.applyEdits(batch);
		}, plan.message);
	}
	render(state: ViewerState): void {
		const panel = this.root.querySelector<HTMLElement>('[data-ribbon-tab="cross-functional"]');
		if (!panel) return;
		const options = swimlaneSelection(state)?.info.options;
		const set = (id: string, command: Command) => {
			const reason = this.reason(state, command);
			const item = panel.querySelector<RibbonCommand>(`[command="${id}"]`);
			if (!item) return;
			item.disabled = reason !== undefined;
			if (reason) item.title = reason;
		};
		set('swimlane-separator', 'separator');
		set('swimlane-rotate-labels', 'rotate-labels');
		set('swimlane-insert-before', 'insert-before');
		set('swimlane-insert-after', 'insert-after');
		set('swimlane-delete', 'delete');
		const menus: [string, Command][] = [
			['swimlane-insert', 'insert-after'],
			['swimlane-style', 'style'],
		];
		for (const [id, command] of menus) {
			const button = panel.querySelector<RibbonCommand>(`[data-menu="${id}"]`);
			if (button) button.disabled = this.reason(state, command) !== undefined;
		}
		const rotate = panel.querySelector<HTMLElement>('[command="swimlane-rotate-labels"]');
		rotate?.setAttribute('pressed', String(options?.rotateLabels ?? false));
		for (const [id, on, command] of [
			['swimlane-title', options?.showTitle ?? true, 'title'],
			['swimlane-separators', options?.showSeparators ?? true, 'separators'],
		] as const) {
			const box = panel.querySelector<HTMLElement & { checked: boolean; disabled: boolean }>(
				`[data-check="${id}"]`,
			);
			if (!box) continue;
			box.checked = on;
			box.disabled = this.reason(state, command) !== undefined;
		}
		for (const item of panel.querySelectorAll<HTMLElement>('[command^="swimlane-style-"]'))
			item.setAttribute(
				'checked',
				String(item.getAttribute('command') === `swimlane-style-${options?.style}`),
			);
	}
}
