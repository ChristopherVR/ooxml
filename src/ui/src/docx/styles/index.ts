import { legacyThemeAliases, legacyThemeBridge } from '../../theme';
import aliases from './aliases.css?raw';
import base from './base.css?raw';
import ribbon from './ribbon.css?raw';
import ribbonButtons from './ribbon-buttons.css?raw';
import ribbonGallery from './ribbon-gallery.css?raw';
import ribbonWidgets from './ribbon-widgets.css?raw';
import ribbonStylesGroup from './ribbon-styles-group.css?raw';
import ribbonPopups from './ribbon-popups.css?raw';
import ribbonPolish from './ribbon-polish.css?raw';
import ribbonHomeMenus from './ribbon-home-menus.css?raw';
import documentCss from './document.css?raw';
import underline from './underline.css?raw';
import print from './print.css?raw';
import titlebar from './chrome-titlebar.css?raw';
import backstage from './backstage.css?raw';
import statusDialogs from './status-dialogs.css?raw';
import chromeDocument from './chrome-document.css?raw';
import pageNavigator from './page-navigator.css?raw';
import keyboardMenu from './keyboard-menu.css?raw';
import smartArt from './smartart.css?raw';
import contentControls from './content-controls.css?raw';
import shapesCss from './shapes.css?raw';

// Legacy custom properties remain inputs, never defaults that hide an inherited Office theme.
// Resolve at the host before mapping onto inner shared controls to avoid CSS variable cycles.
export const themeTokenText =
	legacyThemeAliases(':host', '--dve-', '--_dve-') + ':host{color-scheme:inherit;}';

export const editorStyleText = [
	themeTokenText,
	aliases,
	// Private aliases resolve on the host; legacy compatibility only remaps the inner chrome.
	legacyThemeBridge('.dve-frame', '--_dve-'),
	base,
	ribbon,
	ribbonButtons,
	ribbonGallery,
	ribbonWidgets,
	ribbonStylesGroup,
	ribbonPopups,
	ribbonPolish,
	ribbonHomeMenus,
	documentCss,
	underline,
	print,
	titlebar,
	backstage,
	statusDialogs,
	chromeDocument,
	pageNavigator,
	keyboardMenu,
	smartArt,
	contentControls,
	shapesCss,
].join('\n');
