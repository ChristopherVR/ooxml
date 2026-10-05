/**
 * Folds ribbon groups into dropdown buttons when a panel is too narrow. The folding is the shared
 * `createRibbonOverflow` of ooxml-ui; this supplies Excel's icons and command selector.
 */
import { createRibbonOverflow } from '../../controls';
import { setLargeCaption } from './caption.js';
import { ribbonIcon } from './icons';

/** Icon on a collapsed group's button, by the group's English label. */
const GROUP_ICONS: Record<string, string> = {
	Clipboard: 'paste',
	Font: 'fontColor',
	Alignment: 'alignCenter',
	Number: 'numberFormat',
	Styles: 'cellStyles',
	Cells: 'insertCells',
	Editing: 'find',
	Tables: 'table',
	Illustrations: 'picture',
	Charts: 'chart',
	Links: 'link',
	Comments: 'comment',
	Text: 'textBox',
	Symbols: 'symbol',
	'Page Setup': 'margins',
	'Scale to Fit': 'zoom',
	'Sheet Options': 'gridlines',
	'Function Library': 'fx',
	'Defined Names': 'nameManager',
	'Formula Auditing': 'traceArrow',
	Calculation: 'calculate',
	'Sort & Filter': 'sortFilter',
	'Data Tools': 'dataValidation',
	Proofing: 'spelling',
	Protect: 'protect',
	'Sheet View': 'view',
	'Workbook Views': 'normalView',
	Show: 'gridlines',
	Zoom: 'zoom',
	Window: 'freeze',
};

const overflow = createRibbonOverflow({
	icon: ribbonIcon,
	groupIcons: GROUP_ICONS,
	fallbackIcon: 'select',
	ribbon: '.xve-ribbon',
	command: 'button[data-command]',
	setCaption: setLargeCaption,
	reserveCaptions: true,
});

export const fitPanel = overflow.fitPanel;
export const attachRibbonOverflow = overflow.attach;
export const refitRibbon = overflow.refit;
