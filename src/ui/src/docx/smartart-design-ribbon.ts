// The SmartArt Design tab: Word's contextual tab for a selected SmartArt graphic. Add Shape (after,
// before, above, below), Promote, Demote, Move Up, Move Down and the Text Pane toggle edit the
// nodes through the shared core operations; Change Layout opens the shared picker. Word's
// SmartArt Styles and Change Colors mount the shared galleries and source-backed format commands.
import type { EditorState } from 'prosemirror-state';
import {
	DIAGRAM_ADD_SHAPE_COMMANDS,
	DIAGRAM_ADD_SHAPE_LABEL,
	DIAGRAM_DESIGN_COMMANDS,
	DIAGRAM_DESIGN_GROUPS,
	type DiagramDesignCommandKey,
} from 'ooxml-core/diagram';
import { group, menuSelect, stack, tool } from './ribbon-parts';
import { isTabHidden, panelOf, selectTab, setTabHidden } from './ribbon-tab-api';
import { selectedDiagram } from './smartart-commands';
import { smartArtFormatGallery, syncSmartArtFormatGalleries } from './smartart-format-galleries';

// The commands and labels are the shared SmartArt Design list (`ooxml-core/diagram`), which
// Excel's tab is built from too.
const ADD_SHAPE: Array<[string, string]> = DIAGRAM_ADD_SHAPE_COMMANDS.map((c) => [c.key, c.label]);
type AddKey = 'addAfter' | 'addBefore' | 'addAbove' | 'addBelow';
const label = (key: DiagramDesignCommandKey): string =>
	DIAGRAM_DESIGN_COMMANDS.find((command) => command.key === key)?.label ?? key;

export function buildSmartArtDesignPanel(panel: HTMLElement): void {
	panel.append(
		group(
			DIAGRAM_DESIGN_GROUPS.create,
			menuSelect(
				DIAGRAM_ADD_SHAPE_LABEL,
				'addShape',
				ADD_SHAPE,
				(value) => ({ type: 'smartArt', key: value as AddKey }),
				{ momentary: true },
			),
			tool(
				label('textPane'),
				'textPane',
				{ type: 'smartArt', key: 'textPane' },
				{ large: true, className: 'dve-smartart-text-pane-toggle' },
			),
			stack(
				tool(label('promote'), 'promote', { type: 'smartArt', key: 'promote' }, { inline: true }),
				tool(label('demote'), 'demote', { type: 'smartArt', key: 'demote' }, { inline: true }),
			),
			stack(
				tool(label('moveUp'), 'moveUp', { type: 'smartArt', key: 'moveUp' }, { inline: true }),
				tool(
					label('moveDown'),
					'moveDown',
					{ type: 'smartArt', key: 'moveDown' },
					{ inline: true },
				),
			),
		),
		group(
			DIAGRAM_DESIGN_GROUPS.layouts,
			tool(
				label('changeLayout'),
				'changeLayout',
				{ type: 'smartArt', key: 'changeLayout' },
				{ large: true },
			),
		),
		group('SmartArt Styles', smartArtFormatGallery('style'), smartArtFormatGallery('colors')),
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
	const diagram = selectedDiagram(state)?.diagram;
	syncSmartArtFormatGalleries(
		toolbar,
		!selected || !editable,
		diagram?.colorsId,
		diagram?.quickStyleId,
	);
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
