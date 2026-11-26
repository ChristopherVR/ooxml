import { deArrange } from './arrange';
import { deDesign } from './design';
import { deStyleDialog } from './style-dialog';
import { deSplitCells } from './split-cells';
import { deOptions } from './options';
import { deColumns } from './columns';
import { deMultilevelList } from './multilevel-list';
import { deFontAdvanced } from './font-advanced';
import { deLigatures } from './ligatures';
import { deNavigation } from './navigation';
import { deDropCap } from './drop-cap';
import { deFindStories } from './find-stories';
import { dePasteOptions } from './paste-options';
import { dePicture } from './picture';
import { dePictureWatermark } from './picture-watermark';
import { deConvertNotes } from './convert-notes';
import { deDocumentInspector } from './document-inspector';
import { deLineNumbers } from './line-numbers';
import { deObjects } from './objects';
import { deShapes } from '../../ribbon/shape-names';
import { deSmartArt } from './smartart';
import { deCharts } from './charts';
import type { LocaleStrings } from './en';
import { deApp } from './de-app';
import { deTableProperties } from './table-properties';
import { deRibbon } from './de-ribbon';
import { deRibbonNames } from './ribbon-names';

export const de: LocaleStrings = {
	...deArrange,
	...deDesign,
	...deStyleDialog,
	...deOptions,
	...deRibbon,
	...deRibbonNames,
	...deApp,
	...deTableProperties,
	...deLineNumbers,
	...deDropCap,
	...deFindStories,
	...dePasteOptions,
	...dePicture,
	...dePictureWatermark,
	...deConvertNotes,
	...deDocumentInspector,
	...deNavigation,
	...deColumns,
	...deSplitCells,
	...deMultilevelList,
	...deFontAdvanced,
	...deLigatures,
	...deObjects,
	...deShapes,
	...deSmartArt,
	...deCharts,
};
