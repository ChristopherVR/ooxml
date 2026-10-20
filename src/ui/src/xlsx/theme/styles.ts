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
	':host{color-scheme:inherit;}' + legacyThemeAliases(':host', '--xve-', '--xlsx-theme-');

/** The shell stylesheet (tokens, chrome, ribbon, popups, backstage) for the shadow root. */
export const editorStyleText = [
	themeTokenText,
	legacyThemeBridge('.xve-frame', '--xlsx-theme-'),
	shell,
	ribbon,
	ribbonParity,
	popups,
	backstage,
	chartSeries,
	withTokens(rangeControl),
].join('\n');
