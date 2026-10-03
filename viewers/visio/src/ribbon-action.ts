/** Every command the Visio ribbon, status bar or a shortcut can raise, as a `ribbon-action` event. */
export type VisioRibbonAction =
	| { type: 'history'; key: 'undo' | 'redo' }
	| { type: 'tool'; tool: 'pointer' | 'rectangle' }
	| { type: 'delete' }
	| { type: 'pane'; pane: 'pages' | 'inspector' }
	| { type: 'reveal'; panel: 'edit' | 'notes' | 'selection' | 'layers'; focusText?: boolean }
	| { type: 'grid' }
	| { type: 'fullscreen' }
	| { type: 'zoom'; mode: 'fit' | 'width' | 'actual' }
	| { type: 'search' }
	| { type: 'page'; step: 1 | -1 };

/** Event name shared by every ribbon control; detail is the typed action. */
export const RIBBON_ACTION_EVENT = 'ribbon-action';
export type RibbonActionEvent = CustomEvent<VisioRibbonAction>;

/** Internal, non-composed event: it stays inside the viewer's shadow tree. */
export function emitRibbonAction(target: EventTarget, action: VisioRibbonAction): void {
	target.dispatchEvent(new CustomEvent(RIBBON_ACTION_EVENT, { detail: action, bubbles: true }));
}
