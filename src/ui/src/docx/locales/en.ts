import { enArrange } from './arrange';
import { viewsStrings } from './document-views';
import { enDesign } from './design';
import { enTableTab } from './en-table-tab';
import { enStyleDialog } from './style-dialog';
import { enSplitCells } from './split-cells';
import { enObjects } from './objects';
import { enShapes } from '../../ribbon/shape-names';
import { enSmartArt } from './smartart';
import { enCharts } from './charts';
import { enOptions } from './options';
import { enColumns } from './columns';
import { enMultilevelList } from './multilevel-list';
import { enFontAdvanced } from './font-advanced';
import { enLigatures } from './ligatures';
import { enNavigation } from './navigation';
import { enDropCap } from './drop-cap';
import { enFindStories } from './find-stories';
import { enPasteOptions } from './paste-options';
import { enPicture } from './picture';
import { enPictureWatermark } from './picture-watermark';
import { enConvertNotes } from './convert-notes';
import { enDocumentInspector } from './document-inspector';
import { enLineNumbers } from './line-numbers';
import { enApp } from './en-app';
import { enTableProperties } from './table-properties';
import { enRibbon } from './en-ribbon';
import { enRibbonNames } from './ribbon-names';

/**
 * English display strings, the source of truth for the key set. Keys are the English UI text the
 * localizer looks up (or dotted ids for templates); every other locale must define every key.
 */
export const en = {
	...enArrange,
	...viewsStrings['en'],
	...enTableTab,
	...enDesign,
	...enStyleDialog,
	...enOptions,
	...enRibbon,
	...enRibbonNames,
	...enApp,
	...enTableProperties,
	...enLineNumbers,
	...enDropCap,
	...enFindStories,
	...enPasteOptions,
	...enPicture,
	...enPictureWatermark,
	...enConvertNotes,
	...enDocumentInspector,
	...enNavigation,
	...enColumns,
	...enSplitCells,
	...enMultilevelList,
	...enFontAdvanced,
	...enLigatures,
	...enObjects,
	...enShapes,
	...enSmartArt,
	...enCharts,
} as const;

export type LocalizationKey = keyof typeof en;
/** A complete locale: adding a key to `en` without translating it is a compile error. */
export type LocaleStrings = Record<LocalizationKey, string>;
export type RibbonStrings = Record<keyof typeof enRibbon, string>;
export type AppStrings = Record<keyof typeof enApp, string>;
