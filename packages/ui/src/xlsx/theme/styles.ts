import { darkTheme, lightTheme } from 'ooxml-core/xlsx/ui';
import { themeToCssVars } from 'ooxml-core/xlsx/ui';
import type { XlsxTheme } from 'ooxml-core/xlsx/ui';
import officeBridge from './office-bridge.css?raw';
import shell from './shell.css?raw';
import ribbon from './ribbon.css?raw';
import popups from './popups.css?raw';
import backstage from './backstage.css?raw';

const declarations = (theme: XlsxTheme, scheme: 'light' | 'dark') =>
	Object.entries(themeToCssVars(theme))
		.map(([name, value]) => `${name}:${value};`)
		.join('') + `color-scheme:${scheme};`;

/**
 * Token presets generated from theme/defaults.ts so the TypeScript presets and the shadow-root CSS
 * cannot drift. `theme="auto"` (the default) follows the OS; explicit values always win.
 */
export const themeTokenText = `:host{${declarations(lightTheme, 'light')}}
@media (prefers-color-scheme: dark){:host(:not([theme="light"])){${declarations(darkTheme, 'dark')}}}
:host([theme="dark"]){${declarations(darkTheme, 'dark')}}
:host([theme="light"]){${declarations(lightTheme, 'light')}}`;

/** The shell stylesheet (tokens, chrome, ribbon, popups, backstage) for the shadow root. */
export const editorStyleText = [
	themeTokenText,
	officeBridge,
	shell,
	ribbon,
	popups,
	backstage,
].join('\n');
