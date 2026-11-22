// The Picture Format tab: Word's contextual tab for a selected picture (Arrange > Wrap Text and
// Size > Crop). It appears while a picture is selected and reads the wrapping back from it.
import type { EditorState } from 'prosemirror-state';
import { group, menuSelect, tool } from './ribbon-parts';
import { isTabHidden, panelOf, selectTab, setTabHidden } from './ribbon-tab-api';
import { pictureWrapOf, selectedPicturePos } from './picture-format-commands';
import { WRAPS } from './shape-format-ribbon';

export function buildPictureFormatPanel(panel: HTMLElement): void {
	const graphicsFill = tool(
		'Graphics Fill',
		'formatPicture',
		{ type: 'picture', key: 'graphicsFill' },
		{ large: true },
	);
	graphicsFill.dataset.graphicsFill = '';
	panel.append(
		group('Graphics Styles', graphicsFill),
		group(
			'Arrange',
			menuSelect(
				'Wrap text',
				'columns',
				WRAPS,
				(value) => ({ type: 'picture', key: 'wrap', value }),
				{ momentary: true },
			),
		),
		group(
			'Size',
			tool('Crop', 'formatPicture', { type: 'picture', key: 'crop' }, { large: true }),
			tool(
				'Format picture',
				'formatPicture',
				{ type: 'formatPicture' },
				{ large: true, caption: 'Format', name: 'ribbon.name.formatPicture' },
			),
		),
	);
}

/** Shows the tab while a picture is selected (switching to it on entry) and fills its controls. */
export function syncPictureFormat(
	toolbar: HTMLElement,
	state: EditorState,
	editable: boolean,
): void {
	const panel = panelOf(toolbar, 'picture-format');
	if (!panel) return;
	const pos = selectedPicturePos(state);
	const entering = isTabHidden(panel) && pos !== undefined;
	setTabHidden(toolbar, 'picture-format', 'contextual', pos === undefined);
	if (entering) selectTab(toolbar, 'picture-format');
	for (const control of panel.querySelectorAll<HTMLButtonElement | HTMLSelectElement>(
		'button, select',
	))
		control.disabled = !editable || pos === undefined;
	const node = pos === undefined ? null : state.doc.nodeAt(pos);
	const wrap = panel.querySelector<HTMLSelectElement>(
		'[data-localearialabel="Wrap text"], [aria-label="Wrap text"]',
	);
	if (node && wrap) wrap.value = pictureWrapOf(node);
	const graphics = panel.querySelector<HTMLButtonElement>('[data-graphics-fill]');
	if (graphics) {
		graphics.hidden = !node?.attrs.svgPartName;
		graphics.disabled = !editable || !node?.attrs.svgPartName;
	}
}
