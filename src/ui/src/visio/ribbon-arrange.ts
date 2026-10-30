import { paintOptions } from './ribbon-style-options';
import { colorGrid, type ColorTarget } from './ribbon-color-menu';
import { effectsOptions, quickStyleGallery } from './ribbon-shape-styles';
import { group, menu, stack } from './ribbon-parts';

const ARRANGE = 'Select shapes to align, distribute, rotate or flip them.';

/** Home > Fill or Line: Office's colour picker over the menu's own commands. */
function paintMenu(doc: Document, target: Exclude<ColorTarget, 'font'>, label: string) {
	const el = menu(doc, {
		id: target,
		label,
		icon: target,
		size: 'small',
		items: paintOptions(target),
	});
	el.prepend(colorGrid(doc, target));
	return el;
}

/** Shape Styles: Quick Styles, Fill, Line and Effects (Home and Shape Format). */
export function shapeStylesGroup(doc: Document): HTMLElement {
	return group(
		doc,
		'Shape Styles',
		[
			quickStyleGallery(doc),
			stack(doc, [
				paintMenu(doc, 'fill', 'Fill'),
				paintMenu(doc, 'line', 'Line'),
				menu(doc, {
					id: 'effects',
					label: 'Effects',
					icon: 'effects',
					size: 'small',
					items: effectsOptions(),
				}),
			]),
		],
		{ dialog: 'Format Shape' },
	);
}

/** Arrange: Align, Position, order and Group (Home and Shape Format). */
export function arrangeGroup(doc: Document): HTMLElement {
	return group(doc, 'Arrange', [
		menu(doc, {
			id: 'align',
			label: 'Align',
			icon: 'alignObjects',
			items: (['left', 'center', 'right', 'top', 'middle', 'bottom'] as const).map((edge) => ({
				id: `align-shapes-${edge}`,
				label: `Align ${edge[0]!.toUpperCase()}${edge.slice(1)}`,
				action: { type: 'arrange', operation: { type: 'align', edge } },
			})),
		}),
		menu(doc, {
			id: 'position',
			label: 'Position',
			icon: 'position',
			unsupported: ARRANGE,
			items: [
				{ id: 'auto-align', label: 'Auto Align & Space', action: { type: 'auto-align' } },
				{
					id: 'distribute-horizontal',
					label: 'Distribute Horizontally',
					action: { type: 'arrange', operation: { type: 'distribute', axis: 'horizontal' } },
				},
				{
					id: 'distribute-vertical',
					label: 'Distribute Vertically',
					action: { type: 'arrange', operation: { type: 'distribute', axis: 'vertical' } },
				},
				{
					id: 'rotate',
					label: 'Rotate Shapes',
					items: [
						{
							id: 'rotate-left',
							label: 'Rotate Left 90°',
							action: { type: 'rotate', direction: 'left' },
						},
						{
							id: 'rotate-right',
							label: 'Rotate Right 90°',
							action: { type: 'rotate', direction: 'right' },
						},
						{
							id: 'flip-horizontal',
							label: 'Flip Horizontal',
							action: { type: 'flip', axis: 'horizontal' },
						},
						{
							id: 'flip-vertical',
							label: 'Flip Vertical',
							action: { type: 'flip', axis: 'vertical' },
						},
					],
				},
			],
		}),
		stack(doc, [
			menu(doc, {
				id: 'bring-to-front',
				label: 'Bring to Front',
				icon: 'bringToFront',
				size: 'small',
				split: true,
				action: { type: 'shape-order', order: 'front' },
				items: [
					{
						id: 'bring-forward',
						label: 'Bring Forward',
						action: { type: 'shape-order', order: 'forward' },
					},
				],
			}),
			menu(doc, {
				id: 'send-to-back',
				label: 'Send to Back',
				icon: 'sendToBack',
				size: 'small',
				split: true,
				action: { type: 'shape-order', order: 'back' },
				items: [
					{
						id: 'send-backward',
						label: 'Send Backward',
						action: { type: 'shape-order', order: 'backward' },
					},
				],
			}),
			menu(doc, {
				id: 'group',
				label: 'Group',
				icon: 'group',
				size: 'small',
				items: [
					{
						id: 'group-shapes',
						label: 'Group',
						icon: 'group',
						action: { type: 'grouping', operation: 'group' },
						keys: ['Control+Shift+G', 'Ctrl+Shift+G'],
					},
					{
						id: 'ungroup',
						label: 'Ungroup',
						action: { type: 'grouping', operation: 'ungroup' },
						keys: ['Control+Shift+U', 'Ctrl+Shift+U'],
					},
				],
			}),
		]),
	]);
}
