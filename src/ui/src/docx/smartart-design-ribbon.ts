// The SmartArt Design tab: Word's contextual tab for a selected SmartArt graphic. Add Shape (after,
// before, above, below), Promote, Demote, Move Up, Move Down and the Text Pane toggle edit the
// nodes through the shared core operations; Change Layout opens the shared picker. Word's
// SmartArt Styles and Change Colors galleries are not offered: only one colour and one style
// definition ship with the core, so there is nothing yet to switch between.
import type { EditorState } from 'prosemirror-state';
import { group, menuSelect, stack, tool } from './ribbon-parts';
import { isTabHidden, panelOf, selectTab, setTabHidden } from './ribbon-tab-api';
import { selectedDiagram } from './smartart-commands';

const ADD_SHAPE: Array<[string, string]> = [
	['addAfter', 'Add Shape After'],
	['addBefore', 'Add Shape Before'],
	['addAbove', 'Add Shape Above'],
	['addBelow', 'Add Shape Below'],
];
type AddKey = 'addAfter' | 'addBefore' | 'addAbove' | 'addBelow';

export function buildSmartArtDesignPanel(panel: HTMLElement): void {
	panel.append(
		group(
			'Create Graphic',
			menuSelect(
				'Add Shape',
				'addShape',
				ADD_SHAPE,
				(value) => ({ type: 'smartArt', key: value as AddKey }),
				{ momentary: true },
			),
			tool(
				'Text Pane',
				'textPane',
				{ type: 'smartArt', key: 'textPane' },
				{ large: true, className: 'dve-smartart-text-pane-toggle' },
			),
			stack(
				tool('Promote', 'promote', { type: 'smartArt', key: 'promote' }, { inline: true }),
				tool('Demote', 'demote', { type: 'smartArt', key: 'demote' }, { inline: true }),
			),
			stack(
				tool('Move Up', 'moveUp', { type: 'smartArt', key: 'moveUp' }, { inline: true }),
				tool('Move Down', 'moveDown', { type: 'smartArt', key: 'moveDown' }, { inline: true }),
			),
		),
		group(
			'Layouts',
			tool(
				'Change Layout',
				'changeLayout',
				{ type: 'smartArt', key: 'changeLayout' },
				{ large: true },
			),
		),
	);
}

/**
 * Shows the tab while a SmartArt is selected (switching to it on entry), enables its commands when
 * the document can be edited and reflects the text pane's state on its toggle.
 */
export function syncSmartArtDesign(
	toolbar: HTMLElement,
	state: EditorState,
	editable: boolean,
	textPaneOpen: boolean,
): void {
	const panel = panelOf(toolbar, 'smartart-design');
	if (!panel) return;
	const selected = Boolean(selectedDiagram(state));
	const entering = isTabHidden(panel) && selected;
	setTabHidden(toolbar, 'smartart-design', 'contextual', !selected);
	if (entering) selectTab(toolbar, 'smartart-design');
	for (const control of panel.querySelectorAll<HTMLButtonElement | HTMLSelectElement>(
		'button, select',
	))
		control.disabled = !selected || !editable;
	panel
		.querySelector('.dve-smartart-text-pane-toggle')
		?.setAttribute('aria-pressed', String(textPaneOpen));
}
