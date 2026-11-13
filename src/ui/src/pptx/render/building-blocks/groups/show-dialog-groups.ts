/**
 * The slide show dialogs the building blocks never mount (Set Up Slide Show, Custom Shows,
 * Broadcast, the keyboard shortcut help) as action groups over a small port per binding. Settings
 * and show edits are the shared `building-block-show-settings` and `building-block-custom-shows`
 * decisions; the groups add the session's gating and request completion.
 *
 * @module render/building-blocks/groups/show-dialog-groups
 */
import type { PptxCustomShow, PptxPresentationProperties, PptxSlide } from 'ooxml-core/pptx';

import { BUILDING_BLOCK_OK, buildingBlockFailure } from '../../building-block-actions';
import type { BuildingBlockResult } from '../../building-block-actions';
import {
	applyCustomShowEdit,
	describeCustomShows,
	resolveCustomShowStart,
} from '../../building-block-custom-shows';
import type {
	CustomShowEdit,
	CustomShowSlideChoice,
	CustomShowSummary,
} from '../../building-block-custom-shows';
import {
	describeBroadcastSupport,
	describeShortcutCatalogue,
} from '../../building-block-show-dialogs';
import type { ShortcutCatalogueEntry } from '../../building-block-show-dialogs';
import {
	describeSlideShowSettings,
	resolveSlideShowSettings,
} from '../../building-block-show-settings';
import type {
	SlideShowSettingsPatch,
	SlideShowSettingsView,
} from '../../building-block-show-settings';
import { defineBuildingBlockActionGroup } from '../action-groups';

/** Set Up Slide Show, as a host's own dialog sees it. */
export interface BuildingBlockSetUpSlideShow {
	/** The current settings with every default resolved. */
	current: SlideShowSettingsView;
	/** Store the changed fields (the dialog's OK). Validated; the deck is marked dirty. */
	apply: (patch: SlideShowSettingsPatch) => BuildingBlockResult;
}

/** Custom Shows, as a host's own dialog sees it. Every edit is stored at once. */
export interface BuildingBlockCustomShows {
	shows: CustomShowSummary[];
	/** The deck's slides a show may list, in deck order. */
	slides: CustomShowSlideChoice[];
	/** The show Set Up and the Slide Show tab currently use, or null. */
	activeId: string | null;
	/** Create a show (the active slide when `slideRIds` is omitted is not assumed). */
	create: (name: string, slideRIds?: readonly string[]) => BuildingBlockResult;
	rename: (id: string, name: string) => BuildingBlockResult;
	/** Replace the show's slide list (relationship ids, in show order). */
	setSlides: (id: string, slideRIds: readonly string[]) => BuildingBlockResult;
	copy: (id: string, name?: string) => BuildingBlockResult;
	/** Delete a show; jumps into it are cleared. */
	remove: (id: string) => BuildingBlockResult;
	/** Make `id` the active show and start the slide show from its first slide. */
	start: (id: string) => BuildingBlockResult;
}

export interface BuildingBlockShortcuts {
	/** The shortcut help is open (also toggled by `?` and Ctrl+/). The host renders it. */
	open: boolean;
	setOpen: (open: boolean) => void;
	toggle: () => void;
	/** The keymap as data, in the viewer's current language. */
	catalogue: ShortcutCatalogueEntry[];
}

export interface SetUpSlideShowPort {
	presentationProperties(): PptxPresentationProperties;
	slideCount(): number;
	customShows(): readonly PptxCustomShow[];
	/** Store validated settings and mark the deck dirty (the bundled dialog's OK). */
	save(settings: PptxPresentationProperties): void;
}

const settingsContext = (port: SetUpSlideShowPort) => ({
	slideCount: port.slideCount(),
	customShows: [...port.customShows()],
});

export const setUpSlideShowGroup = defineBuildingBlockActionGroup({
	id: 'setUpSlideShow',
	create: (context, port: SetUpSlideShowPort): BuildingBlockSetUpSlideShow => ({
		current: describeSlideShowSettings(port.presentationProperties(), settingsContext(port)),
		apply: (patch) => {
			const denied = context.refuse('setUpSlideShow');
			if (denied) {
				return denied;
			}
			const resolved = resolveSlideShowSettings(
				port.presentationProperties(),
				patch,
				settingsContext(port),
			);
			if (!resolved.ok) {
				return buildingBlockFailure(resolved.code);
			}
			port.save(resolved.settings);
			return context.complete('setUpSlideShow', BUILDING_BLOCK_OK);
		},
	}),
});

export interface CustomShowsPort {
	/**
	 * The shows including edits stored in this tick: a binding whose state lands later keeps the
	 * last committed list itself, so a rename followed by a slide change builds on the rename.
	 */
	customShows(): readonly PptxCustomShow[];
	slides(): readonly PptxSlide[];
	activeId(): string | null;
	/** Store an edit (the slides too when jumps into a removed show were cleared). */
	commit(shows: PptxCustomShow[], slides: PptxSlide[] | undefined): void;
	/** Make `id` the active show and start the slide show from its beginning. */
	start(id: string): void;
}

export const customShowsGroup = defineBuildingBlockActionGroup({
	id: 'customShows',
	create: (context, port: CustomShowsPort): BuildingBlockCustomShows => {
		const edit = (change: CustomShowEdit): BuildingBlockResult => {
			const denied = context.refuse('customShows');
			if (denied) {
				return denied;
			}
			const outcome = applyCustomShowEdit(port.customShows(), port.slides(), change);
			if (!outcome.ok) {
				return buildingBlockFailure(outcome.code);
			}
			port.commit(outcome.shows, outcome.slides);
			return BUILDING_BLOCK_OK;
		};
		return {
			...describeCustomShows(port.customShows(), port.slides()),
			activeId: port.activeId(),
			create: (name, slideRIds) =>
				edit({ kind: 'create', name, ...(slideRIds ? { slideRIds } : {}) }),
			rename: (id, name) => edit({ kind: 'rename', id, name }),
			setSlides: (id, slideRIds) => edit({ kind: 'setSlides', id, slideRIds }),
			copy: (id, name) => edit({ kind: 'copy', id, ...(name === undefined ? {} : { name }) }),
			remove: (id) => edit({ kind: 'remove', id }),
			start: (id) => {
				if (!context.ready()) {
					return buildingBlockFailure('not-ready');
				}
				const target = resolveCustomShowStart(port.customShows(), port.slides(), id);
				if (!target.ok) {
					return buildingBlockFailure(target.code);
				}
				port.start(id);
				return context.complete('customShows', BUILDING_BLOCK_OK);
			},
		};
	},
});

/** Broadcast cannot be completed by the building blocks; see `BroadcastSupport`. */
export const broadcastGroup = defineBuildingBlockActionGroup({
	id: 'broadcast',
	create: (_context, _port: null) => describeBroadcastSupport(),
});

export interface ShortcutsPort {
	open(): boolean;
	setOpen(open: boolean): void;
	toggle(): void;
	/** The binding's translator, for the catalogue's descriptions and headings. */
	translate(key: string): string;
}

export const shortcutsGroup = defineBuildingBlockActionGroup({
	id: 'shortcuts',
	create: (_context, port: ShortcutsPort): BuildingBlockShortcuts => ({
		open: port.open(),
		setOpen: (open) => port.setOpen(open),
		toggle: () => port.toggle(),
		catalogue: describeShortcutCatalogue((key) => port.translate(key)),
	}),
});
