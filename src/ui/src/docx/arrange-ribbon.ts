import type { EditorState } from 'prosemirror-state';
import { drawingWrapOf } from 'ooxml-core/docx';
import { drawingPositions } from './arrange-selection';
import type { RibbonAction } from './ribbon-action';
import { isDrawingWrap } from './drawing-layout-default';
import { translate, type EditorLocale, type LocalizationKey } from './localization';
import { parsePlacement, shapeSpecOf } from './print-float-locate';
import { menuGallery } from './ribbon-gallery-menu';
import type { RibbonIcon } from './ribbon-icons';
import { group, menuSelect, stack, tool } from './ribbon-parts';
import { textDirectionCellAvailable } from './table-properties';

export const WRAPS: Array<[string, string]> = [
	['inline', 'In Line with Text'],
	['square', 'Square'],
	['tight', 'Tight'],
	['through', 'Through'],
	['topAndBottom', 'Top and Bottom'],
	['behind', 'Behind Text'],
	['front', 'In Front of Text'],
];

const positions: Array<[string, string]> = [
	['inline', 'In Line with Text'],
	['top-left', 'Top Left'],
	['top-center', 'Top Center'],
	['top-right', 'Top Right'],
	['center-left', 'Middle Left'],
	['center-center', 'Middle Center'],
	['center-right', 'Middle Right'],
	['bottom-left', 'Bottom Left'],
	['bottom-center', 'Bottom Center'],
	['bottom-right', 'Bottom Right'],
];

function menu(
	key: string,
	label: string,
	icon: RibbonIcon,
	values: Array<[string, string]>,
	action: (value: string) => RibbonAction = (value) => ({ type: 'arrange', key, value }),
	inline = key !== 'wrap' && key !== 'direction',
) {
	const wrapper = menuSelect(label, icon, values, action, {
		momentary: true,
		inline,
	});
	wrapper.querySelector('select')!.dataset.arrange = key;
	return wrapper;
}

export function textDirectionMenu() {
	return menu('direction', 'Text Direction', 'textBox', [
		['horizontal', 'Horizontal'],
		['clockwise', 'Rotate Text 90°'],
		['counterclockwise', 'Rotate Text 270°'],
	]);
}

function arrangeButton(label: string, icon: RibbonIcon, key: 'pane' | 'alt') {
	const control = tool(label, icon, { type: 'arrange', key, value: '' }, { inline: true });
	control.dataset.arrange = key;
	return control;
}

export function buildArrangeGroup(kind?: 'picture' | 'shape') {
	const position = menuGallery(
		'Position',
		'pageNumber',
		'drawingPosition',
		positions,
		(value) => ({ type: 'arrange', key: 'position', value }),
		{ momentary: true },
	);
	position.classList.add('ribbon-position');
	position.querySelector('select')!.dataset.arrange = 'position';
	const positionNote = document.createElement('small');
	positionNote.className = 'ribbon-position-note';
	positionNote.hidden = true;
	position.append(positionNote);
	return group(
		'Arrange',
		position,
		menu(
			'wrap',
			'Wrap text',
			'columns',
			[
				...WRAPS,
				['move', 'Move with Text'],
				['fix', 'Fix Position on Page'],
				['default', 'Set as Default Layout'],
				['points', 'Edit Wrap Points (unavailable)'],
			],
			(value) =>
				kind && isDrawingWrap(value)
					? { type: kind, key: 'wrap', value }
					: { type: 'arrange', key: 'wrap', value },
		),
		stack(
			menu('order', 'Bring Forward', 'bringForward', [
				['forward', 'Bring Forward'],
				['first', 'Bring to Front'],
				['front', 'In Front of Text'],
			]),
			menu('order', 'Send Backward', 'sendBackward', [
				['backward', 'Send Backward'],
				['last', 'Send to Back'],
				['behind', 'Behind Text'],
			]),
			arrangeButton('Selection Pane', 'selectionPane', 'pane'),
		),
		stack(
			menu('align', 'Align', 'columns', [
				['page', 'Align to Page'],
				['margin', 'Align to Margin'],
				['selected', 'Align Selected Objects'],
				['left', 'Align Left'],
				['center', 'Align Center'],
				['right', 'Align Right'],
				['top', 'Align Top'],
				['middle', 'Align Middle'],
				['bottom', 'Align Bottom'],
				['x', 'Distribute Horizontally'],
				['y', 'Distribute Vertically'],
			]),
			menu('group', 'Group', 'shapes', [
				['group', 'Group'],
				['ungroup', 'Ungroup'],
			]),
			menu('rotate', 'Rotate', 'orientation', [
				['right', 'Rotate Right 90°'],
				['left', 'Rotate Left 90°'],
				['flipH', 'Flip Horizontal'],
				['flipV', 'Flip Vertical'],
				['more', 'More Rotation Options'],
			]),
		),
		stack(arrangeButton('Alt Text', 'textBox', 'alt')),
	);
}

export function syncArrange(
	toolbar: HTMLElement,
	state: EditorState,
	editable: boolean,
	locale: EditorLocale,
	printLayout: boolean,
) {
	const positions = drawingPositions(state);
	const node = positions.length ? state.doc.nodeAt(state.selection.from) : null;
	const drawing = editable && Boolean(node);
	const floated = drawing && Boolean(node?.attrs.anchored);
	const spec = node ? shapeSpecOf(node) : undefined;
	const reason: LocalizationKey = editable ? 'Select a drawing first.' : 'Editing is unavailable.';
	const parent = node && state.doc.resolve(state.selection.from).parent;
	const groupable =
		positions.length >= 2 &&
		positions.length <= 100 &&
		positions.every((pos) => {
			const item = state.doc.nodeAt(pos);
			const shape = item && shapeSpecOf(item);
			return (
				item?.attrs.anchored &&
				shape?.editable &&
				!shape.groupChildren &&
				!shape.text?.length &&
				!item.attrs.textBoxContent &&
				!item.attrs.altText &&
				!item.attrs.title &&
				!item.marks.length &&
				state.doc.resolve(pos).parent === parent
			);
		});
	const ungroupable =
		positions.length === 1 &&
		Boolean(
			spec?.groupCanUngroup &&
			!(spec.rotation || spec.flipH || spec.flipV) &&
			!node?.attrs.altText &&
			!node?.attrs.title,
		);
	const alignable =
		positions.length > 0 &&
		positions.every((pos) => Boolean(state.doc.nodeAt(pos)?.attrs.anchored));
	for (const control of toolbar.querySelectorAll<HTMLSelectElement>('select[data-arrange]')) {
		const key = control.dataset.arrange;
		let unavailable: LocalizationKey | undefined;
		if (!editable) unavailable = 'Editing is unavailable.';
		else if (key === 'direction') {
			if (!node?.attrs.textBoxContent && !textDirectionCellAvailable(state))
				unavailable = 'Select a text box or editable table cell.';
		} else if ((key === 'align' || key === 'group' || key === 'position') && !printLayout)
			unavailable = 'Switch to Print Layout to arrange drawings.';
		else if (!drawing) unavailable = reason;
		else if (key === 'align' && !alignable) unavailable = 'Select a floating drawing.';
		else if (key === 'group' && !(floated && (groupable || ungroupable)))
			unavailable = 'Select two or more floating shapes in one paragraph.';
		const possible = unavailable === undefined;
		control.disabled = !possible;
		const description = unavailable && translate(locale, unavailable);
		control.title = description ?? control.getAttribute('aria-label') ?? '';
		if (description) control.setAttribute('aria-description', description);
		else control.removeAttribute('aria-description');
		if (key === 'position') {
			const note = control.parentElement?.querySelector<HTMLElement>('.ribbon-position-note');
			if (note) {
				note.hidden = printLayout;
				note.textContent = translate(locale, 'Print Layout only');
			}
		}
		const disableOption = (value: string, disabled: boolean, reason?: LocalizationKey) => {
			const option = [...control.options].find((candidate) => candidate.value === value);
			if (option) {
				option.disabled = disabled;
				option.title = disabled && reason ? translate(locale, reason) : '';
			}
		};
		if (key === 'wrap') {
			control.value = node
				? drawingWrapOf(Boolean(node.attrs.anchored), parsePlacement(node.attrs.placement))
				: '';
			const reason = printLayout
				? 'Select a floating drawing.'
				: 'Switch to Print Layout to arrange drawings.';
			disableOption('move', !floated || !printLayout, reason);
			disableOption('fix', !floated || !printLayout, reason);
			disableOption('points', true, 'Wrap point editing is unavailable.');
		}
		if (key === 'order')
			for (const value of ['forward', 'first', 'backward', 'last']) disableOption(value, !floated);
		if (key === 'align')
			for (const [value, count] of [
				['selected', 2],
				['x', 3],
				['y', 3],
			] as const)
				disableOption(
					value,
					positions.length < count,
					value === 'selected' ? 'Select two or more floating drawings.' : undefined,
				);
		if (key === 'group') {
			disableOption('group', !groupable, 'Select two or more floating shapes in one paragraph.');
			disableOption('ungroup', !ungroupable);
		}
	}
	for (const button of toolbar.querySelectorAll<HTMLButtonElement>('button[data-arrange="alt"]')) {
		button.disabled = !drawing;
		button.title = drawing ? (button.getAttribute('aria-label') ?? '') : translate(locale, reason);
	}
	for (const button of toolbar.querySelectorAll<HTMLButtonElement>('button[data-arrange="pane"]'))
		button.disabled = false;
}
