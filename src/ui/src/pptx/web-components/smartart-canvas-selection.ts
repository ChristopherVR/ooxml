import { smartArtNodeAtPoint } from '../render/smartart-3d-edit-layer';

const EDITOR = '[data-pptx-editor-chrome]';
const VIEWPORT = '[data-pptx-viewport]';
const NODE = '[data-smartart-node-id]';
const SELECTED = 'data-pptx-smartart-selected';
const SELECTION_CHANGE = 'pptx-smartart-selection-change';

interface NodeSelection {
	elementId: string;
	nodeIds: ReadonlySet<string>;
}

const selections = new WeakMap<Element, NodeSelection>();

function paintSelection(editor: Element, selection: NodeSelection | undefined): void {
	const viewport = editor.querySelector(VIEWPORT);
	for (const node of viewport?.querySelectorAll(`${NODE},[${SELECTED}]`) ?? []) {
		const id = node.getAttribute('data-smartart-node-id');
		const element = node.closest('[data-element-id]');
		node.toggleAttribute(
			SELECTED,
			Boolean(
				id &&
				selection?.nodeIds.has(id) &&
				element?.getAttribute('data-element-id') === selection.elementId,
			),
		);
	}
}

/** The ribbon tab row stays mounted while the canvas selection changes. */
export function trackSmartArtCanvasSelection(host: HTMLElement, signal: AbortSignal): void {
	const editor = host.closest(EDITOR);
	if (!editor) return;
	editor.addEventListener(
		'pointerdown',
		(input) => {
			const event = input as PointerEvent;
			const target = event.target;
			if (event.button !== 0 || !(target instanceof Element) || target.closest(EDITOR) !== editor)
				return;
			if (!target.closest(VIEWPORT) && !target.closest('[data-pptx-chrome="slides"]')) return;
			const element = target.closest('[data-pptx-element="true"][data-element-id]');
			const node =
				target.closest(NODE) ??
				(element && smartArtNodeAtPoint(element, event.clientX, event.clientY));
			const nodeId = node?.getAttribute('data-smartart-node-id');
			const elementId = element?.getAttribute('data-element-id');
			const previous = selections.get(editor);
			let selection: NodeSelection | undefined;
			if (elementId && nodeId) {
				const multi =
					previous?.elementId === elementId && (event.ctrlKey || event.metaKey || event.shiftKey);
				const ids = new Set(multi ? previous.nodeIds : []);
				if (ids.has(nodeId)) ids.delete(nodeId);
				else ids.add(nodeId);
				selection = { elementId, nodeIds: ids };
				if (multi) {
					// Keep the binding's existing frame selection while toggling its members.
					event.preventDefault();
					event.stopPropagation();
				}
			}
			if (selection) selections.set(editor, selection);
			else selections.delete(editor);
			paintSelection(editor, selection);
			editor.dispatchEvent(new Event(SELECTION_CHANGE));
		},
		{ capture: true, signal },
	);
}

export function watchSmartArtCanvasSelection(
	host: HTMLElement,
	onChange: () => void,
	signal: AbortSignal,
): void {
	host.closest(EDITOR)?.addEventListener(SELECTION_CHANGE, onChange, { signal });
}

export function selectedSmartArtCanvasNodes(
	host: HTMLElement,
	elementId: string,
): readonly string[] {
	const editor = host.closest(EDITOR);
	const selection = editor && selections.get(editor);
	if (!editor || selection?.elementId !== elementId) return [];
	paintSelection(editor, selection);
	return [...selection.nodeIds];
}
