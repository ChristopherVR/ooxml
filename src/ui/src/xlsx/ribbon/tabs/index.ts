// Every ribbon tab UI-COMMANDS supplies, in Excel's order. The File backstage belongs to the shell.
import type { RibbonTab } from 'ooxml-core/xlsx/ui';
import { chartDesignTab, tableDesignTab } from './contextual';
import { dataTab, formulasTab } from './formulas-data';
import { homeTab } from './home';
import { insertTab, pageLayoutTab } from './insert-layout';
import { helpTab, reviewTab, viewTab } from 'ooxml-core/xlsx/ui';

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
