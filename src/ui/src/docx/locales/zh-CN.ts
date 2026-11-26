import { zhEquations } from './equations';
import { zhArrange } from './arrange';
import { viewsStrings } from './document-views';
import { zhDesign } from './design';
import { tableTabBase } from './table-tab-base';
import { zhCNStyleDialog } from './style-dialog';
import { zhSplitCells } from './split-cells';
import { zhOptions } from './options';
import { zhColumns } from './columns';
import { zhMultilevelList } from './multilevel-list';
import { zhFontAdvanced } from './font-advanced';
import { zhLigatures } from './ligatures';
import { zhNavigation } from './navigation';
import { zhLineNumbers } from './line-numbers';
import { zhDropCap } from './drop-cap';
import { zhFindStories } from './find-stories';
import { zhPasteOptions } from './paste-options';
import { zhPicture } from './picture';
import { zhPictureWatermark } from './picture-watermark';
import { zhConvertNotes } from './convert-notes';
import { zhDocumentInspector } from './document-inspector';
import { zhObjects } from './objects';
import { zhShapes } from '../../ribbon/shape-names';
import { zhSmartArt } from './smartart';
import { zhCharts } from './charts';
import type { BaseLocaleStrings } from './en';
import { zhCNApp } from './zh-CN-app';
import { zhTableProperties } from './table-properties';
import { zhCNRibbon } from './zh-CN-ribbon';
import { zhCNRibbonNames } from './ribbon-names';

export const zhCN: BaseLocaleStrings = {
	...zhEquations,
	...zhArrange,
	...viewsStrings['zh-CN'],
	...tableTabBase['zh-CN'],
	...zhDesign,
	...zhCNStyleDialog,
	...zhOptions,
	...zhCNRibbon,
	...zhCNRibbonNames,
	...zhCNApp,
	...zhTableProperties,
	...zhLineNumbers,
	...zhDropCap,
	...zhFindStories,
	...zhPasteOptions,
	...zhPicture,
	...zhPictureWatermark,
	...zhConvertNotes,
	...zhDocumentInspector,
	...zhNavigation,
	...zhColumns,
	...zhSplitCells,
	...zhMultilevelList,
	...zhFontAdvanced,
	...zhLigatures,
	...zhObjects,
	...zhShapes,
	...zhSmartArt,
	...zhCharts,
};
