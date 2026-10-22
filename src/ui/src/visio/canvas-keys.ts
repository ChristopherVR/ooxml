import { editErrorMessage, isEditCancellation } from 'ooxml-core/visio/ui';
import { canEditDiagram, type ViewerController } from './controller';
import type { CanvasEditing } from './canvas-editing';
import { planDeleteSelection } from './viewer-delete';

/** Run a diagram edit, reporting a failure other than a cancellation. */
async function attempt(run: () => Promise<void>, announce: (m: string) => void, done?: string) {
	try {
		await run();
		if (done) announce(done);
	} catch (error) {
		if (!isEditCancellation(error)) announce(editErrorMessage(error));
	}
}
/** Delete the selection through the controller; false when there is nothing deletable. */
export function deleteDiagramSelection(
	controller: ViewerController,
	announce: (message: string) => void,
): Promise<void> {
	const state = controller.state;
	const plan = planDeleteSelection(state);
	if (!plan || !canEditDiagram(state) || state.edit.busy || state.loading) return Promise.resolve();
	return attempt(() => controller.applyEdits(plan.edits), announce, plan.message);
}
/**
 * The few keys a bare drawing surface answers, gated like the full viewer's: Delete removes the
 * selection, Escape drops the draw tool, Ctrl+Z and Ctrl+Y walk the history.
 */
export function wireCanvasKeys(
	viewport: HTMLElement,
	controller: ViewerController,
	editing: CanvasEditing,
	announce: (message: string) => void,
): () => void {
	const onKey = (event: KeyboardEvent) => {
		const typing = event
			.composedPath()
			.some(
				(node) =>
					node instanceof Element && node.matches('input, textarea, select, [contenteditable]'),
			);
		if (typing || event.isComposing || editing.edit.editing) return;
		const state = controller.state;
		const control = event.ctrlKey || event.metaKey;
		if (event.key === 'Escape' && state.tool !== 'pointer') {
			editing.cancelTools();
			controller.setTool('pointer');
		} else if ((event.key === 'Delete' || event.key === 'Backspace') && !control) {
			if (!state.selectedShapes.length) return;
			event.preventDefault();
			void deleteDiagramSelection(controller, announce);
		} else if (control && !event.altKey && /^[zy]$/i.test(event.key)) {
			if (!canEditDiagram(state) || state.edit.busy || state.loading) return;
			const redo = event.key.toLowerCase() === 'y' || event.shiftKey;
			if (redo ? !state.edit.canRedo : !state.edit.canUndo) return;
			event.preventDefault();
			void attempt(() => (redo ? controller.redo() : controller.undo()), announce);
		}
	};
	viewport.addEventListener('keydown', onKey);
	return () => viewport.removeEventListener('keydown', onKey);
}
