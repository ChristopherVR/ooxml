// Every ribbon tab UI-COMMANDS supplies, in Excel's order. The File backstage belongs to the shell.
import type { RibbonTab } from 'ooxml-core/xlsx/ui';
import {
	chartDesignTab,
	embeddedObjectTab,
	graphicsFormatTab,
	slicerTab,
	smartArtDesignTab,
	tableDesignTab,
	timelineTab,
} from './contextual';
import { dataTab, formulasTab } from './formulas-data';
import { homeTab } from './home';
import { drawTab } from './draw';
import { chartFormatTab } from './chart-format';
import { pictureFormatTab, shapeFormatTab } from './drawing-format';
import { insertTab, pageLayoutTab } from './insert-layout';
import { helpTab, reviewTab, viewTab } from 'ooxml-core/xlsx/ui';
import { sparklineTab } from './sparkline';
import { pivotAnalyzeTab, pivotDesignTab } from './pivot';

export function commandTabs(): RibbonTab[] {
	return [
		homeTab(),
		insertTab(),
		drawTab(),
		pageLayoutTab(),
		formulasTab(),
		dataTab(),
		reviewTab(),
		viewTab(),
		helpTab(),
		tableDesignTab(),
		chartDesignTab(),
		chartFormatTab(),
		pictureFormatTab(),
		shapeFormatTab(),
		graphicsFormatTab(),
		embeddedObjectTab(),
		smartArtDesignTab(),
		slicerTab(),
		timelineTab(),
		sparklineTab(),
		pivotAnalyzeTab(),
		pivotDesignTab(),
	];
}
