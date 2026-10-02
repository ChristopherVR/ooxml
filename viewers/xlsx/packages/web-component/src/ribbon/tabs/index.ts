// Every ribbon tab UI-COMMANDS supplies, in Excel's order. The File backstage belongs to the shell.
import type { RibbonTab } from '../parts.js';
import { chartDesignTab, tableDesignTab } from './contextual.js';
import { dataTab, formulasTab } from './formulas-data.js';
import { homeTab } from './home.js';
import { insertTab, pageLayoutTab } from './insert-layout.js';
import { helpTab, reviewTab, viewTab } from './review-view.js';

export function commandTabs(): RibbonTab[] {
	return [
		homeTab(),
		insertTab(),
		pageLayoutTab(),
		formulasTab(),
		dataTab(),
		reviewTab(),
		viewTab(),
		helpTab(),
		tableDesignTab(),
		chartDesignTab(),
	];
}
