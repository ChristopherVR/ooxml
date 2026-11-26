import type { PptxElement } from 'ooxml-core/pptx';

import type { AlignEdge } from './element-align';
import { canInteractWithElement } from './element-locks';
import type { RibbonHomeItem } from './ribbon-home-spec';

const ALIGN = {
	left: ['pptx.ribbon.alignLeft', 'Align Left'],
	centerH: ['pptx.ribbon.alignCenter', 'Align Center'],
	right: ['pptx.ribbon.alignRight', 'Align Right'],
	top: ['pptx.ribbon.alignTop', 'Align Top'],
	middle: ['pptx.ribbon.alignMiddle', 'Align Middle'],
	bottom: ['pptx.ribbon.alignBottom', 'Align Bottom'],
} as const;

export interface DrawingArrangeActions {
	align?: ((edge: AlignEdge) => void) | undefined;
	distribute?: ((axis: 'horizontal' | 'vertical') => void) | undefined;
	flip?: ((axis: 'horizontal' | 'vertical') => void) | undefined;
	rotation?:
		| (() => {
				element: PptxElement | null | undefined;
				patch: (id: string, patch: { rotation: number }) => void;
		  })
		| undefined;
	group?: (() => void) | undefined;
	ungroup?: (() => void) | undefined;
	position?: (() => void) | undefined;
	selectionPane?: (() => void) | undefined;
}

interface ArrangeItem extends RibbonHomeItem {
	action?: keyof DrawingArrangeActions;
}

const item = (
	value: string,
	labelKey: string,
	fallback: string,
	control: string,
	action?: keyof DrawingArrangeActions,
	separator = false,
): ArrangeItem => ({
	value,
	labelKey,
	fallback,
	...(action && { action }),
	separator,
	attrs: { 'data-ribbon-control': control },
	descriptionKey: 'pptx.insert.needsSelection',
	description: 'Select an object first.',
});
const legacy = (name: string) => `home.arrange.${name}`;

export const HOME_DRAWING_ARRANGE_ITEMS: readonly ArrangeItem[] = [
	item('forward', 'pptx.contextMenu.bringForward', 'Bring Forward', legacy('bringForward')),
	item('backward', 'pptx.contextMenu.sendBackward', 'Send Backward', legacy('sendBackward')),
	item('front', 'pptx.contextMenu.bringToFront', 'Bring to Front', legacy('bringToFront')),
	item('back', 'pptx.contextMenu.sendToBack', 'Send to Back', legacy('sendToBack')),
	item('group', 'pptx.ribbon.group', 'Group', legacy('group'), 'group', true),
	item('ungroup', 'pptx.ribbon.ungroup', 'Ungroup', legacy('ungroup'), 'ungroup'),
	item(
		'position',
		'pptx.arrange.positionSize',
		'Position & Size',
		'home.drawing.positionSize',
		'position',
		true,
	),
	...Object.entries(ALIGN).map(([edge, [key, label]]) =>
		item(edge, key, label, legacy('align'), 'align', edge === 'left'),
	),
	item(
		'distribute-horizontal',
		'pptx.arrange.distributeHorizontal',
		'Distribute Horizontally',
		legacy('align'),
		'distribute',
	),
	item(
		'distribute-vertical',
		'pptx.arrange.distributeVertical',
		'Distribute Vertically',
		legacy('align'),
		'distribute',
	),
	item(
		'rotate-right',
		'pptx.pictureFormat.rotateRight',
		'Rotate Right 90°',
		'home.drawing.rotateRight',
		'rotation',
		true,
	),
	item(
		'rotate-left',
		'pptx.pictureFormat.rotateLeft',
		'Rotate Left 90°',
		'home.drawing.rotateLeft',
		'rotation',
	),
	item(
		'flip-horizontal',
		'pptx.arrange.flipHorizontally',
		'Flip Horizontally',
		legacy('flipHorizontal'),
		'flip',
	),
	item(
		'flip-vertical',
		'pptx.arrange.flipVertically',
		'Flip Vertically',
		legacy('flipVertical'),
		'flip',
	),
	item(
		'selectionPane',
		'pptx.selectionPane.title',
		'Selection Pane',
		'home.drawing.selectionPane',
		'selectionPane',
		true,
	),
];

export interface DrawingArrangeInput {
	actions: DrawingArrangeActions;
	canEdit: boolean;
	selected: boolean;
	canGroup: boolean;
	canDistribute: boolean;
	canUngroup: boolean;
}

export function drawingArrangeItems(input?: DrawingArrangeInput): RibbonHomeItem[] {
	const rotationTarget = input?.actions.rotation?.().element;
	return HOME_DRAWING_ARRANGE_ITEMS.filter(
		({ action }) => !action || Boolean(input?.actions[action]),
	).map(({ action, ...row }) => {
		if (!input) return row;
		switch (action) {
			case 'group':
				return { ...row, disabled: !input.canEdit || !input.canGroup };
			case 'ungroup':
				return { ...row, disabled: !input.canEdit || !input.canUngroup };
			case 'distribute':
				return { ...row, disabled: !input.canEdit || !input.canDistribute };
			case 'position':
				return { ...row, disabled: !input.selected };
			case 'rotation':
				return {
					...row,
					disabled:
						!input.canEdit || !rotationTarget || !canInteractWithElement(rotationTarget, 'rotate'),
				};
			case 'selectionPane':
				return row;
			default:
				return { ...row, disabled: !input.canEdit || !input.selected };
		}
	});
}

export function runDrawingArrange(value: string, actions: DrawingArrangeActions): boolean {
	switch (value) {
		case 'left':
		case 'centerH':
		case 'right':
		case 'top':
		case 'middle':
		case 'bottom':
			actions.align?.(value);
			return true;
		case 'distribute-horizontal':
		case 'distribute-vertical':
			actions.distribute?.(value === 'distribute-horizontal' ? 'horizontal' : 'vertical');
			return true;
		case 'flip-horizontal':
		case 'flip-vertical':
			actions.flip?.(value === 'flip-horizontal' ? 'horizontal' : 'vertical');
			return true;
		case 'rotate-right':
		case 'rotate-left': {
			const target = actions.rotation?.();
			const element = target?.element;
			if (target && element && canInteractWithElement(element, 'rotate')) {
				target.patch(element.id, quarterTurnPatch(element, value === 'rotate-right'));
			}
			return true;
		}
		case 'group':
			actions.group?.();
			return true;
		case 'ungroup':
			actions.ungroup?.();
			return true;
		case 'position':
			actions.position?.();
			return true;
		case 'selectionPane':
			actions.selectionPane?.();
			return true;
		default:
			return false;
	}
}

export function quarterTurnPatch(element: PptxElement, clockwise: boolean): { rotation: number } {
	const delta = clockwise ? 90 : -90;
	return { rotation: ((((element.rotation ?? 0) + delta) % 360) + 360) % 360 };
}
