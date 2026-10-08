import { visioDuplicateCommand, visioSelectionIsOnPage } from 'ooxml-core/visio/ui';
import type { ViewerController, ViewerState } from './controller';
import type { RibbonCommand } from './ribbon-parts';

/** Source duplication availability; the controller owns the transaction and new selection. */
export class ViewerDuplication {
	constructor(
		private readonly root: ShadowRoot,
		private readonly controller: ViewerController,
		private readonly edit: (run: () => Promise<void>, message: string) => void,
	) {}
	private reason(state: ViewerState): string {
		if (!state.edit.sourceAvailable)
			return 'Open a .vsdx file to duplicate shapes. Model-only documents are read only.';
		if (state.loading || state.edit.busy) return 'Wait for the current operation to finish.';
		if (!state.selectedShapes.length) return 'Select shapes to duplicate.';
		const page = state.document?.pages[state.pageIndex];
		if (!page || state.selectedShapes.some((shape) => !visioSelectionIsOnPage(shape, page.id)))
			return 'Select shapes on the current page to duplicate.';
		if (
			!visioDuplicateCommand(
				page,
				state.selectedShapes.map((shape) => shape.id),
			)
		)
			return 'Duplicate requires ordinary local two-dimensional shapes without masters, groups, or connections. Source protection and formulas are checked when applying the edit.';
		return '';
	}
	run(): void {
		if (this.reason(this.controller.state)) return;
		this.edit(() => this.controller.duplicateSelection(), 'Duplicated selected shapes.');
	}
	render(state: ViewerState): void {
		const reason = this.reason(state);
		for (const id of ['duplicate', 'ctx-duplicate']) {
			const command = this.root.querySelector<RibbonCommand>(`[command="${id}"]`);
			if (!command) continue;
			command.disabled = !!reason;
			command.title = reason ? `Duplicate (Ctrl+D): ${reason}` : 'Duplicate (Ctrl+D)';
		}
		const paste = this.root.querySelector<RibbonCommand>('[data-menu="paste"]');
		if (paste) paste.disabled = !!reason;
	}
}
