// The Shape Format tab: Word's contextual tab for a selected drawing shape (fill, outline, text,
// wrapping and size). It appears while a shape is selected and reads its values back from it.
import type { EditorState } from 'prosemirror-state';
import { colorSplit, group, menuSelect, spinner, stack, tool } from './ribbon-parts';
import { panelOf, selectTab, setTabHidden, isTabHidden } from './ribbon-tab-api';
import { selectedShape, shapeWrapOf } from './shape-commands';

const WRAPS: Array<[string, string]> = [
	['inline', 'In Line with Text'],
	['square', 'Square'],
	['tight', 'Tight'],
	['through', 'Through'],
	['topAndBottom', 'Top and Bottom'],
	['behind', 'Behind Text'],
	['front', 'In Front of Text'],
];
const WEIGHTS: Array<[string, string]> = [
	['0.25', '0.25 pt'],
	['0.5', '0.5 pt'],
	['0.75', '0.75 pt'],
	['1', '1 pt'],
	['1.5', '1.5 pt'],
	['2.25', '2.25 pt'],
	['3', '3 pt'],
	['4.5', '4.5 pt'],
	['6', '6 pt'],
];
const PX_PER_INCH = 96;

export function buildShapeFormatPanel(panel: HTMLElement): void {
	panel.append(
		group(
			'Shape Styles',
			colorSplit(
				'Shape fill',
				'shading',
				[],
				(value) => ({ type: 'shape', key: 'fill', value }),
				'#4472c4',
				{ noneLabel: 'No Fill' },
			),
			colorSplit(
				'Shape outline',
				'borders',
				[],
				(value) => ({ type: 'shape', key: 'outline', value }),
				'#2f528f',
				{ noneLabel: 'No Outline' },
			),
			menuSelect(
				'Shape outline weight',
				'lineSpacing',
				WEIGHTS,
				(value) => ({ type: 'shape', key: 'weight', value: Number(value) }),
				{ momentary: true },
			),
		),
		group(
			'Text',
			tool(
				'Add text',
				'textBox',
				{ type: 'formatDialog', kind: 'shapeText' },
				{ large: true, caption: 'Add Text' },
			),
		),
		group(
			'Arrange',
			menuSelect(
				'Wrap text',
				'columns',
				WRAPS,
				(value) => ({ type: 'shape', key: 'wrap', value }),
				{ momentary: true },
			),
		),
		group(
			'Size',
			stack(
				spinner('Shape height (inches)', (value) => ({ type: 'shape', key: 'height', value }), {
					min: 0.05,
					max: 22,
					step: 0.01,
				}),
				spinner('Shape width (inches)', (value) => ({ type: 'shape', key: 'width', value }), {
					min: 0.05,
					max: 22,
					step: 0.01,
				}),
			),
		),
	);
}

/** Shows the tab while a shape is selected (switching to it on entry) and fills its controls. */
export function syncShapeFormat(toolbar: HTMLElement, state: EditorState, editable: boolean): void {
	const panel = panelOf(toolbar, 'shape-format');
	if (!panel) return;
	const shape = selectedShape(state);
	const entering = isTabHidden(panel) && !!shape;
	setTabHidden(toolbar, 'shape-format', 'contextual', !shape);
	if (entering) selectTab(toolbar, 'shape-format');
	const active = shape && editable && shape.spec.editable;
	for (const control of panel.querySelectorAll<
		HTMLButtonElement | HTMLSelectElement | HTMLInputElement
	>('button, select, input'))
		control.disabled =
			!active &&
			!(
				control instanceof HTMLButtonElement &&
				control.dataset.action?.includes('shapeText') &&
				shape
			);
	if (!shape) return;
	const set = (label: string, value: string) => {
		const input = panel.querySelector<HTMLInputElement | HTMLSelectElement>(
			`[data-localearialabel="${label}"], [aria-label="${label}"]`,
		);
		const focused =
			input?.getRootNode() instanceof ShadowRoot
				? (input.getRootNode() as ShadowRoot).activeElement
				: document.activeElement;
		if (input && focused !== input) input.value = value;
	};
	const inches = (px: number) => String(+(px / PX_PER_INCH).toFixed(2));
	set('Shape width (inches)', inches(shape.widthPx));
	set('Shape height (inches)', inches(shape.heightPx));
	set('Wrap text', shapeWrapOf(shape));
	const points = (shape.spec.line?.widthEmu ?? 12700) / 12700;
	set('Shape outline weight', WEIGHTS.some(([v]) => Number(v) === points) ? String(points) : '1');
}
