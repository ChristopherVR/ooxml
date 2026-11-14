import { esSplitCells } from './split-cells';
import { esOptions } from './options';
import { esColumns } from './columns';
import { esMultilevelList } from './multilevel-list';
import { esFontAdvanced } from './font-advanced';
import { esLigatures } from './ligatures';
import { esNavigation } from './navigation';
import { esDropCap } from './drop-cap';
import { esFindStories } from './find-stories';
import { esPasteOptions } from './paste-options';
import { esPicture } from './picture';
import { esLineNumbers } from './line-numbers';
import { esObjects } from './objects';
import { esShapes } from './shapes';
import { esSmartArt } from './smartart';
import { esCharts } from './charts';
import type { LocaleStrings } from './en';
import { esApp } from './es-app';
import { esTableProperties } from './table-properties';
import { esRibbon } from './es-ribbon';

export const es: LocaleStrings = {
	...esOptions,
	...esRibbon,
	...esApp,
	...esTableProperties,
	...esLineNumbers,
	...esDropCap,
	...esFindStories,
	...esPasteOptions,
	...esPicture,
	...esNavigation,
	...esColumns,
	...esSplitCells,
	...esMultilevelList,
	...esFontAdvanced,
	...esLigatures,
	...esObjects,
	...esShapes,
	...esSmartArt,
	...esCharts,
};
