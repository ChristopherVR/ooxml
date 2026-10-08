import { darkTheme, lightTheme } from 'ooxml-core/xlsx/ui';
import { themeToCssVars } from 'ooxml-core/xlsx/ui';
import type { XlsxTheme } from 'ooxml-core/xlsx/ui';
import { shadcnBridge } from '../../theme-bridge';
import shell from './shell.css?raw';
import ribbon from './ribbon.css?raw';
import popups from './popups.css?raw';
import backstage from './backstage.css?raw';
import chartSeries from './chart-series.css?raw';
import rangeControl from '../../form/range-control.css?raw';
import { withTokens } from '../../base';

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
	// The shared elements read --office-*; feed them from the editor's --xve-* theme.
	shadcnBridge(':host', '--xve-', {
		extra: { '--office-font': "'Segoe UI', system-ui, -apple-system, sans-serif" },
	}),
	shell,
	ribbon,
	popups,
	backstage,
	chartSeries,
	withTokens(rangeControl),
].join('\n');
