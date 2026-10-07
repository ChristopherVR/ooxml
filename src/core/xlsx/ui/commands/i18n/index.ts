// Every UI-COMMANDS translation table, flattened per locale for `locales/<lang>/commands.ts`.
import { CORE_FEATURE_STRINGS } from './core-features';
import { CHART_SERIES_STRINGS } from './chart-series';
import { FORMAT_CELLS_STRINGS } from './format-cells';
import { MESSAGE_STRINGS } from './messages';
import { NAVIGATION_STRINGS } from './navigation';
import { RIBBON_HOME_STRINGS } from './ribbon-home';
import { RULES_STRINGS } from './rules';
import { RIBBON_TAB_STRINGS } from './ribbon-tabs';
import { TOOLS_STRINGS } from './tools';
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
];

export const commandStrings = (locale: CommandLocale): Record<string, string> =>
	flatten(COMMAND_TABLES, locale);
