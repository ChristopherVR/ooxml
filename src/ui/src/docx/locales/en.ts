import { enObjects } from './objects';
import { enShapes } from './shapes';
import { enCharts } from './charts';
import { enChartFormat } from './chart-format';
import { enOptions } from './options';
import { enColumns } from './columns';
import { enMultilevelList } from './multilevel-list';
import { enFontAdvanced } from './font-advanced';
import { enLigatures } from './ligatures';
import { enNavigation } from './navigation';
import { enDropCap } from './drop-cap';
import { enLineNumbers } from './line-numbers';
import { enApp } from './en-app';
import { enTableProperties } from './table-properties';
import { enRibbon } from './en-ribbon';

/**
 * English display strings, the source of truth for the key set. Keys are the English UI text the
 * localizer looks up (or dotted ids for templates); every other locale must define every key.
 */
export const en = {
	...enOptions,
	...enRibbon,
	...enApp,
	...enTableProperties,
	...enLineNumbers,
	...enDropCap,
	...enNavigation,
	...enColumns,
	...enMultilevelList,
	...enFontAdvanced,
	...enLigatures,
	...enObjects,
	...enShapes,
	...enCharts,
	...enChartFormat,
} as const;

export type LocalizationKey = keyof typeof en;
/** A complete locale: adding a key to `en` without translating it is a compile error. */
export type LocaleStrings = Record<LocalizationKey, string>;
export type RibbonStrings = Record<keyof typeof enRibbon, string>;
export type AppStrings = Record<keyof typeof enApp, string>;
