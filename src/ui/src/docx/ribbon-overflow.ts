import { createRibbonOverflow } from '../controls';
import { localeOf, translate } from './localization';
import { ribbonIcon, type RibbonIcon } from './ribbon-icons';

/** The icon shown on a collapsed group's button, by the group's English name. */
export const GROUP_ICONS: Readonly<Record<string, RibbonIcon>> = {
	Clipboard: 'paste',
	'Drawing Tools': 'addPen',
	Stencils: 'ruler',
	Convert: 'inkToShape',
	Replay: 'inkReplay',
	Font: 'changeCase',
	Paragraph: 'alignLeft',
	Styles: 'moreStyles',
	Editing: 'find',
	Tables: 'table',
	Illustrations: 'picture',
	Links: 'link',
	'Add-ins': 'addIns',
	Pages: 'pageBreak',
	Breaks: 'sectionNext',
	Media: 'picture',
	'Header & Footer': 'header',
	Text: 'dateTime',
	Symbols: 'symbol',
	'Page setup': 'margins',
	'Page numbers': 'pageNumber',
	Indent: 'indent',
	Arrange: 'bringForward',
	Proofing: 'spelling',
	Tracking: 'track',
	Changes: 'accept',
	Comments: 'comment',
	Language: 'spelling',
	'Table of contents': 'toc',
	Captions: 'caption',
	Footnotes: 'footnote',
	Show: 'view',
	Zoom: 'zoom',
	Views: 'view',
	Page: 'pageSize',
	Table: 'table',
	Properties: 'table',
	'Cell Size': 'rowAbove',
	'Cell Alignment': 'alignCenter',
	'Table Data': 'table',
};

/** Detaches whichever collapsed-group panel is open. */

/**
 * Folds groups into dropdown buttons when a panel is too narrow, as Word does when a window
 * narrows. The folding is the shared `createRibbonOverflow` of ooxml-ui; this supplies Word's icons.
 */
const overflow = createRibbonOverflow({
	icon: (_doc, name, size) => ribbonIcon(name as RibbonIcon, size),
	// Read per sync from the panel's own locale, so a language change reaches the chevron.
	moreLabel: (panel) => translate(localeOf(panel), 'More'),
	groupIcons: GROUP_ICONS,
	fallbackIcon: 'moreStyles',
	ribbon: '.dve-ribbon',
	command: 'button[data-action]',
	naturalWidth: true,
	preferExpanded: (group, panel) =>
		(group.dataset.label === 'Add-ins' && panel.dataset.panel === 'Insert') ||
		(panel.clientWidth >= 1600 && group.dataset.label === 'Arrange'),
	// Table Layout folds as Word's does: Data and Alignment first, then Cell Size and the row and
	// column commands, and the Select group last.
	groupPriority: (group) => {
		if (group.parentElement?.dataset.ribbonTab !== 'table-layout') return 0;
		if (group.dataset.label === 'Cell Size') return 1;
		if (group.dataset.label !== 'Table') return 0;
		return group.querySelector('[data-table-command="select"]') ? 3 : 2;
	},
});

export const fitPanel = overflow.fitPanel;
export const attachRibbonOverflow = overflow.attach;
export const closeRibbonOverflow = overflow.close;
export const refitRibbon = overflow.refit;
