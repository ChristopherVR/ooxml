import { frEquations } from './equations';
import { frArrange } from './arrange';
import { viewsStrings } from './document-views';
import { frDesign } from './design';
import { tableTabBase } from './table-tab-base';
import { frStyleDialog } from './style-dialog';
import { frSplitCells } from './split-cells';
import { frOptions } from './options';
import { frColumns } from './columns';
import { frMultilevelList } from './multilevel-list';
import { frFontAdvanced } from './font-advanced';
import { frLigatures } from './ligatures';
import { frNavigation } from './navigation';
import { frDropCap } from './drop-cap';
import { frFindStories } from './find-stories';
import { frPasteOptions } from './paste-options';
import { frPicture } from './picture';
import { frPictureWatermark } from './picture-watermark';
import { frConvertNotes } from './convert-notes';
import { frDocumentInspector } from './document-inspector';
import { frLineNumbers } from './line-numbers';
import { frObjects } from './objects';
import { frShapes } from '../../ribbon/shape-names';
import { frSmartArt } from './smartart';
import { frCharts } from './charts';
import type { BaseLocaleStrings } from './en';
import { frApp } from './fr-app';
import { frTableProperties } from './table-properties';
import { frRibbon } from './fr-ribbon';
import { frRibbonNames } from './ribbon-names';

export const fr: BaseLocaleStrings = {
	...frEquations,
	...frArrange,
	...viewsStrings['fr'],
	...tableTabBase['fr'],
	...frDesign,
	...frStyleDialog,
	...frOptions,
	...frRibbon,
	...frRibbonNames,
	...frApp,
	...frTableProperties,
	...frLineNumbers,
	...frDropCap,
	...frFindStories,
	...frPasteOptions,
	...frPicture,
	...frPictureWatermark,
	...frConvertNotes,
	...frDocumentInspector,
	...frNavigation,
	...frColumns,
	...frSplitCells,
	...frMultilevelList,
	...frFontAdvanced,
	...frLigatures,
	...frObjects,
	...frShapes,
	...frSmartArt,
	...frCharts,
};
