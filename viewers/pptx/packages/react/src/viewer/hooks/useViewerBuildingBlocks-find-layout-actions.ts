import { BUILDING_BLOCK_OK, buildingBlockFailure, resetSlideLayoutPath } from 'ooxml-ui/pptx';
import type { BuildingBlockResult, FindReplaceOptions } from 'ooxml-ui/pptx';
import type { PptxLayoutOption, PptxSlide } from 'pptx-viewer-core';

import type { useFindReplace } from './useFindReplace';
import type {
	BuildingBlockFindReplace,
	BuildingBlockLayoutActions,
} from './useViewerBuildingBlocks-actions-types';

/** Refuses an action (the failure to return) or lets it run (undefined). */
export type RefuseReplace = () => BuildingBlockResult | undefined;

/** Find / Replace as the building blocks expose it, over the viewer's own `useFindReplace`. */
export function buildFindReplaceActions(
	findReplace: ReturnType<typeof useFindReplace>,
	refuseReplace: RefuseReplace,
): BuildingBlockFindReplace {
	const matches = findReplace.findResults;
	const currentIndex = findReplace.findResultIndex;
	const options: FindReplaceOptions = {
		matchCase: findReplace.findMatchCase,
		includeNotes: findReplace.findIncludeNotes,
		includeChartTitles: findReplace.findIncludeChartTitles,
		includeChartData: findReplace.findIncludeChartData,
	};
	return {
		open: findReplace.findReplaceOpen,
		setOpen: findReplace.setFindReplaceOpen,
		query: findReplace.findQuery,
		replacement: findReplace.replaceQuery,
		options,
		matches,
		currentIndex,
		currentMatch: matches[currentIndex] ?? null,
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
		next: () => findReplace.navigateFindResult(1),
		previous: () => findReplace.navigateFindResult(-1),
		replaceCurrent: () =>
			refuseReplace() ??
			(matches[currentIndex] === undefined
				? buildingBlockFailure('no-match')
				: (findReplace.handleReplace(), BUILDING_BLOCK_OK)),
		replaceAll: () =>
			refuseReplace() ??
			(matches.length === 0
				? buildingBlockFailure('no-match')
				: (findReplace.handleReplaceAll(), BUILDING_BLOCK_OK)),
	};
}

export interface LayoutActionDeps {
	activeSlide: PptxSlide | undefined;
	scopedLayouts: PptxLayoutOption[];
	/** Read at call time: the latest layouts the active slide may switch to. */
	getScopedLayouts: () => PptxLayoutOption[];
	getActiveSlide: () => PptxSlide | undefined;
	refuse: (id: 'applyLayout' | 'resetLayout') => BuildingBlockResult | undefined;
	applyLayout: (path: string) => Promise<void>;
}

/** Layout apply / reset with the slide master's layouts as the allowed set. */
export function buildLayoutActions(deps: LayoutActionDeps): {
	layout: BuildingBlockLayoutActions;
	apply: (path: string, requireListed: boolean) => Promise<BuildingBlockResult>;
	reset: () => Promise<BuildingBlockResult>;
} {
	const apply = async (path: string, requireListed: boolean): Promise<BuildingBlockResult> => {
		const denied = deps.refuse('applyLayout');
		if (denied) {
			return denied;
		}
		if (requireListed && !deps.getScopedLayouts().some((layout) => layout.path === path)) {
			return buildingBlockFailure('unknown-layout');
		}
		try {
			await deps.applyLayout(path);
			return BUILDING_BLOCK_OK;
		} catch {
			return buildingBlockFailure('unknown-layout');
		}
	};
	const reset = async (): Promise<BuildingBlockResult> => {
		const path = resetSlideLayoutPath(deps.getActiveSlide());
		if (path === undefined) {
			return buildingBlockFailure('no-layout');
		}
		return deps.refuse('resetLayout') ?? apply(path, false);
	};
	return {
		layout: {
			current: deps.scopedLayouts.find((layout) => layout.path === deps.activeSlide?.layoutPath),
			currentPath: deps.activeSlide?.layoutPath,
			available: deps.scopedLayouts,
			apply: (path) => apply(path, true),
			reset,
		},
		apply,
		reset,
	};
}
