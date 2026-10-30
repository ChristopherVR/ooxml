// Every UI-COMMANDS translation table, flattened per locale for `locales/<lang>/commands.ts`.
import { AUDIT_STRINGS } from './audit';
import { CORE_FEATURE_STRINGS } from './core-features';
import { CHART_SERIES_STRINGS } from './chart-series';
import { FILTER_STRINGS } from './filters';
import { FORMAT_CELLS_STRINGS } from './format-cells';
import { MESSAGE_STRINGS } from './messages';
import { NAVIGATION_STRINGS } from './navigation';
import { RIBBON_HOME_STRINGS } from './ribbon-home';
import { RULES_STRINGS } from './rules';
import { RIBBON_TAB_STRINGS } from './ribbon-tabs';
import { TOOLS_STRINGS } from './tools';
import { WHAT_IF_STRINGS } from './what-if';
import { SPARKLINE_STRINGS } from './sparklines';
import { PIVOT_STRINGS } from './pivots';
import { SLICER_STRINGS } from './slicers';
import { PRINT_LAYOUT_STRINGS } from './print-layout';
import { SMARTART_STRINGS } from './smartart';
import { type CommandLocale, type Translations, flatten } from './types';

export const COMMAND_TABLES: readonly Translations[] = [
	RIBBON_HOME_STRINGS,
	RIBBON_TAB_STRINGS,
	MESSAGE_STRINGS,
	NAVIGATION_STRINGS,
	TOOLS_STRINGS,
	FORMAT_CELLS_STRINGS,
	RULES_STRINGS,
	CORE_FEATURE_STRINGS,
	CHART_SERIES_STRINGS,
	SPARKLINE_STRINGS,
	PIVOT_STRINGS,
	SLICER_STRINGS,
	PRINT_LAYOUT_STRINGS,
	WHAT_IF_STRINGS,
	SMARTART_STRINGS,
	FILTER_STRINGS,
	AUDIT_STRINGS,
];

export const commandStrings = (locale: CommandLocale): Record<string, string> =>
	flatten(COMMAND_TABLES, locale);
