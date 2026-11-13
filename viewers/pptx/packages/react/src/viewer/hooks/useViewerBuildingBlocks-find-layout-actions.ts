import type { BuildingBlockFindReplacePort } from 'ooxml-ui/pptx';

import type { useFindReplace } from './useFindReplace';

/** The shared find port over the viewer's own `useFindReplace`. */
export function findReplacePort(
	findReplace: ReturnType<typeof useFindReplace>,
): BuildingBlockFindReplacePort {
	return {
		open: findReplace.findReplaceOpen,
		setOpen: findReplace.setFindReplaceOpen,
		query: findReplace.findQuery,
		replacement: findReplace.replaceQuery,
		options: {
			matchCase: findReplace.findMatchCase,
			includeNotes: findReplace.findIncludeNotes,
			includeChartTitles: findReplace.findIncludeChartTitles,
			includeChartData: findReplace.findIncludeChartData,
		},
		matches: findReplace.findResults,
		currentIndex: findReplace.findResultIndex,
		setQuery: findReplace.setFindQuery,
		setReplacement: findReplace.setReplaceQuery,
		setOptions: (change) => {
			if (change.matchCase !== undefined) {
				findReplace.setFindMatchCase(change.matchCase);
			}
			if (change.includeNotes !== undefined) {
				findReplace.setFindIncludeNotes(change.includeNotes);
			}
			if (change.includeChartTitles !== undefined) {
				findReplace.setFindIncludeChartTitles(change.includeChartTitles);
			}
			if (change.includeChartData !== undefined) {
				findReplace.setFindIncludeChartData(change.includeChartData);
			}
		},
		find: findReplace.performFind,
		navigate: (delta) => findReplace.navigateFindResult(delta),
		replaceCurrent: findReplace.handleReplace,
		replaceAll: findReplace.handleReplaceAll,
	};
}
