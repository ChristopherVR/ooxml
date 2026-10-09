import type { VisioShapeFormatEdit } from 'ooxml-core/visio';
import type { VisioArrangement } from 'ooxml-core/visio/ui';

export type CanvasTool = 'pointer' | 'rectangle' | 'ellipse' | 'line' | 'text';
export type VisioFormattingAction =
	| { type: 'text-toggle'; property: 'bold' | 'italic' | 'underline' | 'strikethrough' }
	| { type: 'font-color'; value?: string }
	| { type: 'text-bullets' }
	| { type: 'text-indent'; direction: 'increase' | 'decrease' }
	| { type: 'font-family'; value: string }
	| { type: 'font-size'; value: number }
	| { type: 'font-step'; direction: 1 | -1 }
	| { type: 'text-align'; axis: 'horizontal'; value: 'left' | 'center' | 'right' | 'justify' }
	| { type: 'text-align'; axis: 'vertical'; value: 'top' | 'middle' | 'bottom' }
	| { type: 'shape-format'; patch: Omit<VisioShapeFormatEdit, 'type' | 'pageId' | 'shapeId'> }
	| { type: 'shape-order'; order: 'front' | 'back' | 'forward' | 'backward' };
/** Every command the Visio ribbon, status bar or a shortcut can raise, as a `ribbon-action` event. */
export type VisioRibbonAction =
	| VisioFormattingAction
	| { type: 'arrange'; operation: VisioArrangement }
	| { type: 'selection'; mode: 'all' | 'clear' }
	| { type: 'history'; key: 'undo' | 'redo' }
	| { type: 'tool'; tool: CanvasTool }
	| { type: 'cancel-drawing' }
	| { type: 'page-insert' }
	| { type: 'paint-properties' }
	| { type: 'delete' }
	| { type: 'duplicate' }
	| { type: 'clipboard'; operation: 'copy' | 'cut' | 'paste'; event?: ClipboardEvent }
	| { type: 'rotate'; direction: 'left' | 'right' }
	| { type: 'flip'; axis: 'horizontal' | 'vertical' }
	| { type: 'pane'; pane: 'shapes' | 'inspector' }
	| { type: 'reveal'; panel: 'edit' | 'notes' | 'selection' | 'layers'; focusText?: boolean }
	| { type: 'grid' }
	| { type: 'panZoom' }
	| { type: 'sizePosition' }
	| { type: 'ruler' }
	| { type: 'fullscreen' }
	| { type: 'presentation' }
	| { type: 'zoom'; mode: 'fit' | 'width' | 'actual' }
	| { type: 'zoomTo'; percent: number }
	| { type: 'search' }
	| { type: 'replace' }
	| { type: 'page'; step: 1 | -1 };

/** Event name shared by every ribbon control; detail is the typed action. */
export const RIBBON_ACTION_EVENT = 'ribbon-action';
export type RibbonActionEvent = CustomEvent<VisioRibbonAction>;

/** Internal, non-composed event: it stays inside the viewer's shadow tree. */
export function emitRibbonAction(target: EventTarget, action: VisioRibbonAction): void {
	target.dispatchEvent(new CustomEvent(RIBBON_ACTION_EVENT, { detail: action, bubbles: true }));
}
