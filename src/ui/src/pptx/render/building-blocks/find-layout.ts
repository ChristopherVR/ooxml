/**
 * Find / Replace and Layout apply / reset as the building blocks expose them, over a small port
 * each binding implements with its own find bar state and layout switching. DOM-free and
 * framework-free.
 *
 * @module render/building-blocks/find-layout
 */
import type { PptxLayoutOption, PptxSlide } from 'ooxml-core/pptx';

import { BUILDING_BLOCK_OK, buildingBlockFailure } from '../building-block-actions';
import type { BuildingBlockResult } from '../building-block-actions';
import type { FindResult } from '../find-replace';
import type { FindReplaceOptions } from '../find-replace-session';
import { resetSlideLayoutPath } from '../ribbon-home-commands';
import type { BuildingBlockFindReplace, BuildingBlockLayoutActions } from './actions-types';

/** A binding's find bar: its current values and the operations behind them. */
export interface BuildingBlockFindReplacePort {
	open: boolean;
	setOpen(open: boolean): void;
	query: string;
	replacement: string;
	options: FindReplaceOptions;
	matches: readonly FindResult[];
	currentIndex: number;
	setQuery(query: string): void;
	setReplacement(replacement: string): void;
	setOptions(change: Partial<FindReplaceOptions>): void;
	find(): void;
	/** Go to the match `delta` (1 or -1) away, wrapping around. */
	navigate(delta: 1 | -1): void;
	replaceCurrent(): void;
	replaceAll(): void;
}

/** Refuses an action (the failure to return) or lets it run (undefined). */
export type RefuseReplace = () => BuildingBlockResult | undefined;

/** Find / Replace as the building blocks expose it. */
export function buildFindReplaceActions(
	port: BuildingBlockFindReplacePort,
	refuseReplace: RefuseReplace,
): BuildingBlockFindReplace {
	const { matches, currentIndex } = port;
	return {
		open: port.open,
		setOpen: port.setOpen,
		query: port.query,
		replacement: port.replacement,
		options: port.options,
		matches,
		currentIndex,
		currentMatch: matches[currentIndex] ?? null,
		setQuery: port.setQuery,
		setReplacement: port.setReplacement,
		setOptions: port.setOptions,
		find: port.find,
		next: () => port.navigate(1),
		previous: () => port.navigate(-1),
		replaceCurrent: () =>
			refuseReplace() ??
			(matches[currentIndex] === undefined
				? buildingBlockFailure('no-match')
				: (port.replaceCurrent(), BUILDING_BLOCK_OK)),
		replaceAll: () =>
			refuseReplace() ??
			(matches.length === 0
				? buildingBlockFailure('no-match')
				: (port.replaceAll(), BUILDING_BLOCK_OK)),
	};
}

/** One element of the active slide holding find matches, for a canvas highlight. */
export interface FindHighlightTarget {
	elementId: string;
	/** The current match is in this element. */
	current: boolean;
}

/**
 * The elements of slide `slideIndex` a canvas outlines for the host find bar's matches, the one
 * with the current match marked. Ports whose text renderer has no run-level highlight (vanilla,
 * and the others until theirs does) outline these elements.
 */
export function findHighlightTargets(
	matches: readonly FindResult[],
	currentIndex: number,
	slideIndex: number,
): FindHighlightTarget[] {
	const targets = new Map<string, boolean>();
	matches.forEach((match, index) => {
		if (match.slideIndex === slideIndex) {
			targets.set(
				match.elementId,
				(targets.get(match.elementId) ?? false) || index === currentIndex,
			);
		}
	});
	return [...targets].map(([elementId, current]) => ({ elementId, current }));
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
