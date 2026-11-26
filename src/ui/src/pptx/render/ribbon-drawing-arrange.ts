import type { PptxElement } from 'ooxml-core/pptx';

import type { RibbonControlId } from './customization/ribbon-control-ids';
import type { AlignEdge } from './element-align';
import { canInteractWithElement } from './element-locks';
import type { RibbonHomeItem } from './ribbon-home-spec';

const ALIGN = {
	left: ['pptx.ribbon.alignLeft', 'Align Left', 'home.arrange.align-left'],
	centerH: ['pptx.ribbon.alignCenter', 'Align Center', 'home.arrange.align-center'],
	right: ['pptx.ribbon.alignRight', 'Align Right', 'home.arrange.align-right'],
	top: ['pptx.ribbon.alignTop', 'Align Top', 'home.arrange.align-top'],
	middle: ['pptx.ribbon.alignMiddle', 'Align Middle', 'home.arrange.align-middle'],
	bottom: ['pptx.ribbon.alignBottom', 'Align Bottom', 'home.arrange.align-bottom'],
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
	attrs: { 'data-ribbon-control': RibbonControlId };
}

const item = (
	value: string,
	labelKey: string,
	fallback: string,
	control: RibbonControlId,
	action?: keyof DrawingArrangeActions,
	separator = false,
): ArrangeItem => ({
	value,
	labelKey,
	fallback,
	...(action && { action }),
	separator,
	attrs: { 'data-ribbon-control': control },
});
const legacy = <Name extends string>(name: Name): `home.arrange.${Name}` => `home.arrange.${name}`;

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
	...Object.entries(ALIGN).map(([edge, [key, label, id]]) =>
		item(edge, key, label, id, 'align', edge === 'left'),
	),
	item(
		'distribute-horizontal',
		'pptx.arrange.distributeHorizontal',
		'Distribute Horizontally',
		legacy('distribute-horizontal'),
		'distribute',
	),
	item(
		'distribute-vertical',
		'pptx.arrange.distributeVertical',
		'Distribute Vertically',
		legacy('distribute-vertical'),
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
	selectedCount: number;
	selectedElement: PptxElement | null | undefined;
	canGroup: boolean;
}

const NEEDS_OBJECT = ['pptx.insert.needsSelection', 'Select an object first.'] as const;
const LOCKED = ['pptx.arrange.objectLocked', 'A selected object is locked.'] as const;

function disabledReason(
	action: ArrangeItem['action'],
	input: DrawingArrangeInput,
	rotationTarget: PptxElement | null | undefined,
): readonly [key: string, text: string] | undefined {
	if (action === 'selectionPane') return;
	if (action === 'position') return input.selectedCount === 0 ? NEEDS_OBJECT : undefined;
	if (!input.canEdit) return ['pptx.draw.readOnly', 'Editing is unavailable.'];

	switch (action) {
		case 'group':
			if (input.selectedCount < 2)
				return ['pptx.arrange.needsGroupSelection', 'Select two or more objects to group.'];
			return input.canGroup ? undefined : LOCKED;
		case 'ungroup':
			if (input.selectedCount !== 1 || input.selectedElement?.type !== 'group')
				return ['pptx.arrange.needsUngroupSelection', 'Select a grouped object to ungroup.'];
			return canInteractWithElement(input.selectedElement, 'group') ? undefined : LOCKED;
		case 'distribute':
			return input.selectedCount < 3
				? ['pptx.arrange.needsDistributeSelection', 'Select three or more objects to distribute.']
				: undefined;
		case 'rotation':
			if (!rotationTarget)
				return ['pptx.arrange.needsRotateSelection', 'Select an object to rotate.'];
			return canInteractWithElement(rotationTarget, 'rotate') ? undefined : LOCKED;
		default:
			return input.selectedCount === 0 ? NEEDS_OBJECT : undefined;
	}
}

export function drawingArrangeItems(input?: DrawingArrangeInput): RibbonHomeItem[] {
	const rotationTarget = input?.actions.rotation?.().element;
	return HOME_DRAWING_ARRANGE_ITEMS.filter(
		({ action }) => !action || Boolean(input?.actions[action]),
	).map(({ action, ...row }) => {
		if (!input) return row;
		const reason = disabledReason(action, input, rotationTarget);
		return reason
			? { ...row, disabled: true, descriptionKey: reason[0], description: reason[1] }
			: row;
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
