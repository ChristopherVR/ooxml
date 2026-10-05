// Every UI-COMMANDS translation table, flattened per locale for `locales/<lang>/commands.ts`.
import { CORE_FEATURE_STRINGS } from './core-features.js';
import { FORMAT_CELLS_STRINGS } from './format-cells.js';
import { MESSAGE_STRINGS } from './messages.js';
import { NAVIGATION_STRINGS } from './navigation.js';
import { RIBBON_HOME_STRINGS } from './ribbon-home.js';
import { RULES_STRINGS } from './rules.js';
import { RIBBON_TAB_STRINGS } from './ribbon-tabs.js';
import { TOOLS_STRINGS } from './tools.js';
import { type CommandLocale, type Translations, flatten } from './types.js';

export const COMMAND_TABLES: readonly Translations[] = [
	RIBBON_HOME_STRINGS,
	RIBBON_TAB_STRINGS,
	MESSAGE_STRINGS,
	NAVIGATION_STRINGS,
	TOOLS_STRINGS,
	FORMAT_CELLS_STRINGS,
	RULES_STRINGS,
	CORE_FEATURE_STRINGS,
];

export const commandStrings = (locale: CommandLocale): Record<string, string> =>
	flatten(COMMAND_TABLES, locale);
