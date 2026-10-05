import { createRibbonOverflow } from 'ooxml-ui/controls';
import { ribbonIcon, type RibbonIcon } from './ribbon-icons';

/** The icon shown on a collapsed group's button, by the group's English name. */
const GROUP_ICONS: Record<string, RibbonIcon> = {
	Clipboard: 'paste',
	Font: 'changeCase',
	Paragraph: 'alignLeft',
	Styles: 'moreStyles',
	Editing: 'find',
	Tables: 'table',
	Illustrations: 'picture',
	Links: 'link',
	Breaks: 'pageBreak',
	'Header & Footer': 'header',
	Text: 'dateTime',
	Symbols: 'symbol',
	'Page setup': 'margins',
	'Page numbers': 'pageNumber',
	'Section breaks': 'sectionNext',
	Indent: 'indent',
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
};

/** Detaches whichever collapsed-group panel is open. */

/**
 * Folds groups into dropdown buttons when a panel is too narrow, as Word does when a window
 * narrows. The folding is the shared `createRibbonOverflow` of ooxml-ui; this supplies Word's icons.
 */
const overflow = createRibbonOverflow({
	icon: (_doc, name, size) => ribbonIcon(name as RibbonIcon, size),
	groupIcons: GROUP_ICONS,
	fallbackIcon: 'moreStyles',
	ribbon: '.dve-ribbon',
	command: 'button[data-action]',
	naturalWidth: true,
});

export const fitPanel = overflow.fitPanel;
export const attachRibbonOverflow = overflow.attach;
export const refitRibbon = overflow.refit;
