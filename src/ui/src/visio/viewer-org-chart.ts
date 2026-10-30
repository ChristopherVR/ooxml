import {
	visioOrgAddPosition,
	visioOrgDeletePosition,
	visioOrgHierarchy,
	visioOrgPosition,
	visioOrgReLayout,
	visioSelectionIsOnPage,
} from 'ooxml-core/visio/ui';
import { canEditDiagram, type ViewerController, type ViewerState } from './controller';
import type { VisioOrgChartAction } from './ribbon-action';
import type { RibbonCommand } from './ribbon-parts';

type Command = VisioOrgChartAction['command'];

/** The org chart position the first selected shape is, on the displayed page. */
export function orgChartSelection(state: ViewerState): { pageId: string; id: string } | undefined {
	const page = state.document?.pages[state.pageIndex];
	const first = state.selectedShapes[0];
	if (!page || !first || !visioSelectionIsOnPage(first, page.id)) return undefined;
	return visioOrgPosition(page, first.id) ? { pageId: page.id, id: first.id } : undefined;
}

/** The Org Chart tab's controller: one edit batch per command, so undo steps back one command. */
export class ViewerOrgChart {
	constructor(
		private readonly root: ShadowRoot,
		private readonly controller: ViewerController,
		private readonly edit: (run: () => Promise<void>, message: string) => void,
	) {}
	reason(state: ViewerState, command: Command): string | undefined {
		if (!canEditDiagram(state)) return 'Open a .vsdx file to edit the org chart.';
		if (state.loading || state.edit.busy) return 'Wait for the current operation to finish.';
		const selection = orgChartSelection(state);
		if (!selection) return 'Select a position in an org chart.';
		const page = state.document!.pages[state.pageIndex]!;
		if (command === 'coworker' && !visioOrgHierarchy(page).parentOf(selection.id))
			return 'The top position has no manager to add a coworker under.';
		return undefined;
	}
	run(action: VisioOrgChartAction): void {
		const state = this.controller.state;
		if (this.reason(state, action.command) !== undefined) return;
		const page = state.document?.pages[state.pageIndex];
		const selection = orgChartSelection(state);
		if (!page || !selection) return;
		const { id } = selection;
		if (action.command === 'delete') {
			const batches = visioOrgDeletePosition(page, id);
			if (!batches) return;
			this.edit(async () => {
				for (const batch of batches) await this.controller.applyEdits(batch);
				this.controller.selectShapes([]);
			}, 'Deleted the position and its connectors.');
		} else if (action.command === 'layout') {
			const edits = visioOrgReLayout(page);
			if (!edits) return;
			this.edit(() => this.controller.applyEdits(edits), 'Laid out the org chart.');
		} else {
			const kind = { subordinate: 'subordinate', coworker: 'peer', assistant: 'assistant' }[
				action.command
			] as 'subordinate' | 'peer' | 'assistant';
			const plan = visioOrgAddPosition(page, id, kind);
			if (!plan) return;
			this.edit(async () => {
				await this.controller.applyEdits(plan.edits);
				this.controller.selectShapes([{ id: plan.newId, name: '', pageId: page.id }]);
			}, `Added a ${action.command}.`);
		}
	}
	render(state: ViewerState): void {
		const panel = this.root.querySelector<HTMLElement>('[data-ribbon-tab="org-chart"]');
		if (!panel) return;
		for (const [id, command] of [
			['org-subordinate', 'subordinate'],
			['org-coworker', 'coworker'],
			['org-assistant', 'assistant'],
			['org-delete', 'delete'],
			['org-layout', 'layout'],
		] as const) {
			const item = panel.querySelector<RibbonCommand>(`[command="${id}"]`);
			if (!item) continue;
			const reason = this.reason(state, command);
			item.disabled = reason !== undefined;
			if (reason) item.title = reason;
		}
	}
}
