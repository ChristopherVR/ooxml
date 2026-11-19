import { legacyThemeAliases, legacyThemeBridge } from '../../theme-bridge';
import shell from './shell.css?raw';
import ribbon from './ribbon.css?raw';
import ribbonParity from './ribbon-parity.css?raw';
import popups from './popups.css?raw';
import backstage from './backstage.css?raw';
import chartSeries from './chart-series.css?raw';
import rangeControl from '../../form/range-control.css?raw';
import { withTokens } from '../../base';

/** Host-computed aliases keep legacy overrides compatible without masking inherited Office tokens. */
export const themeTokenText =
	':host{color-scheme:var(--office-color-scheme, light);}' +
	legacyThemeAliases(':host', '--xve-', '--xlsx-theme-');

/**
 * The shell stylesheet (tokens, chrome, ribbon, popups, backstage) for the shadow root. Bare
 * `var(--office-x)` reads get the token's default, so derived tokens resolve where they are read.
 */
export const editorStyleText = withTokens(
	[
		themeTokenText,
		legacyThemeBridge('.xve-frame', '--xlsx-theme-'),
		shell,
		ribbon,
		ribbonParity,
		popups,
		backstage,
		chartSeries,
		rangeControl,
	].join('\n'),
);
