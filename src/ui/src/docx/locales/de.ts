import { deOptions } from './options';
import { deColumns } from './columns';
import { deMultilevelList } from './multilevel-list';
import { deFontAdvanced } from './font-advanced';
import { deLigatures } from './ligatures';
import { deNavigation } from './navigation';
import { deDropCap } from './drop-cap';
import { deLineNumbers } from './line-numbers';
import { deObjects } from './objects';
import { deShapes } from './shapes';
import { deCharts } from './charts';
import type { LocaleStrings } from './en';
import { deApp } from './de-app';
import { deTableProperties } from './table-properties';
import { deRibbon } from './de-ribbon';

export const de: LocaleStrings = {
	...deOptions,
	...deRibbon,
	...deApp,
	...deTableProperties,
	...deLineNumbers,
	...deDropCap,
	...deNavigation,
	...deColumns,
	...deMultilevelList,
	...deFontAdvanced,
	...deLigatures,
	...deObjects,
	...deShapes,
	...deCharts,
};
