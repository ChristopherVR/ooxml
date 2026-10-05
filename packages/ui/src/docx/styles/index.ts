import { shadcnBridge } from '../../theme';
import { darkTheme, lightTheme } from '../theme/defaults';
import { themeToCssVars } from '../theme/css-vars';
import type { EditorTheme } from '../theme/types';
import aliases from './aliases.css?raw';
import base from './base.css?raw';
import ribbon from './ribbon.css?raw';
import ribbonButtons from './ribbon-buttons.css?raw';
import ribbonGallery from './ribbon-gallery.css?raw';
import ribbonWidgets from './ribbon-widgets.css?raw';
import ribbonStylesGroup from './ribbon-styles-group.css?raw';
import ribbonPopups from './ribbon-popups.css?raw';
import ribbonPolish from './ribbon-polish.css?raw';
import documentCss from './document.css?raw';
import print from './print.css?raw';
import titlebar from './chrome-titlebar.css?raw';
import backstage from './backstage.css?raw';
import statusDialogs from './status-dialogs.css?raw';
import chromeDocument from './chrome-document.css?raw';
import pageNavigator from './page-navigator.css?raw';
import keyboardMenu from './keyboard-menu.css?raw';
import smartArt from './smartart.css?raw';

const declarations = (theme: EditorTheme, scheme: 'light' | 'dark') =>
	Object.entries(themeToCssVars(theme))
		.map(([name, value]) => `${name}:${value};`)
		.join('') + `color-scheme:${scheme};`;

/**
 * Token presets are generated from theme/defaults.ts so the TypeScript presets and the shadow-root
 * CSS cannot drift. `theme="auto"` (the default) follows the OS; explicit values always win.
 */
export const themeTokenText = `:host{${declarations(lightTheme, 'light')}}
@media (prefers-color-scheme: dark){:host(:not([theme="light"])){${declarations(darkTheme, 'dark')}}}
:host([theme="dark"]){${declarations(darkTheme, 'dark')}}
:host([theme="light"]){${declarations(lightTheme, 'light')}}`;

export const editorStyleText = [
	themeTokenText,
	aliases,
	// The shared elements read --office-*; feed them from the editor's --dve-* theme.
	shadcnBridge(':host', '--dve-', { extra: { '--office-font': "'Segoe UI', Arial, sans-serif" } }),
	base,
	ribbon,
	ribbonButtons,
	ribbonGallery,
	ribbonWidgets,
	ribbonStylesGroup,
	ribbonPopups,
	ribbonPolish,
	documentCss,
	print,
	titlebar,
	backstage,
	statusDialogs,
	chromeDocument,
	pageNavigator,
	keyboardMenu,
	smartArt,
].join('\n');
