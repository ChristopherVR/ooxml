/**
 * Normalise a `ViewerCustomization` into the sets every decision function
 * reads, folding the cross-cutting rules in ONCE so no binding re-derives them:
 *
 *  - switching off the `ai` feature removes the AI context-menu entries and the
 *    AI Options page;
 *  - switching off `collaboration` removes the Share and Broadcast dialogs;
 *  - switching off `comments` removes the Add Comment menu entry;
 *  - switching off `presentMode` removes the Slide Show tab;
 *  - switching off `editPoints` removes the Edit Points menu entry;
 *  - hiding a dialog removes every entry point that opens it (toolbar button,
 *    File tab page, action card);
 *  - hiding every export format removes the Export page and button;
 *  - `capabilities` expand into the dialogs, features and commands they gate,
 *    and hiding a dialog or panel removes the buttons, menu entries and
 *    shortcuts that open it (the tables are in `customization-entry-points.ts`).
 *
 * @module render/customization/customization-resolve
 */
import type { BackstagePage } from '../backstage';
import type { BackstageCardId } from '../backstage-cards';
import type { CanvasContextMenuCommandId } from '../canvas-context-menu-commands';
import type { ContextMenuCommandId } from '../context-menu-commands';
import type { EditPointsCommandId } from '../edit-points/edit-points-menu';
import type { EditorKeyActionName } from '../editor-keymap';
import type { RibbonContextualTabId, ToolbarActionId } from '../toolbar-actions';
import { VIEWER_EXPORT_FORMAT_IDS } from './customization-catalog';
import {
	DIALOG_ENTRY_POINTS,
	FEATURE_ENTRY_POINTS,
	PANEL_ENTRY_POINTS,
	expandCapabilities,
} from './customization-entry-points';
import type { EntryPoints } from './customization-entry-points';
import { resolveKeyboardCustomization } from './customization-keymap';
import type { ResolvedKeyboardCustomization } from './customization-keymap';
import type {
	HostBackstagePage,
	ViewerRegionId,
	ViewerRegionRenderer,
} from './customization-regions';
import { resolveRegions } from './customization-resolve-regions';
import { resolveRibbonCustomization } from './customization-resolve-ribbon';
import type {
	ViewerCapabilities,
	OptionsPageId,
	OptionsSectionId,
	OptionsSettingId,
	OptionsSettingValues,
	ViewerCustomization,
	ViewerDialogId,
	ViewerDrawingToolId,
	ViewerFeatureId,
	ViewerPanelId,
} from './customization-types';
import type { RibbonControlId, RibbonGroupId } from './ribbon-control-ids';

/** The normalised, set-based view of a customisation. Treat as immutable. */
export interface ResolvedCustomization {
	hiddenInspectorCards: ReadonlySet<import('./inspector-customization').InspectorCardId>;
	hiddenActions: ReadonlySet<ToolbarActionId>;
	/** Contextual tabs (`shapeFormat`, ...) the host never wants shown. */
	hiddenContextualTabs: ReadonlySet<RibbonContextualTabId>;
	/** Catalogued ribbon groups to hide (`ribbon.hiddenGroups`). */
	hiddenRibbonGroups: ReadonlySet<RibbonGroupId>;
	/** Catalogued ribbon controls to hide (the `<tab>.<group>.<control>` part of `ribbon.hiddenButtons`). */
	hiddenRibbonControls: ReadonlySet<RibbonControlId>;
	hiddenOptionsPages: ReadonlySet<OptionsPageId>;
	hiddenOptionsSections: ReadonlySet<OptionsSectionId>;
	hiddenSettings: ReadonlySet<OptionsSettingId>;
	lockedSettings: OptionsSettingValues;
	defaultSettings: OptionsSettingValues;
	hiddenBackstagePages: ReadonlySet<BackstagePage>;
	hiddenBackstageCards: ReadonlySet<BackstageCardId>;
	hiddenElementCommands: ReadonlySet<ContextMenuCommandId>;
	extraElementCommands: NonNullable<ViewerCustomization['contextMenu']>['extraElementCommands'];
	extraCanvasCommands: NonNullable<ViewerCustomization['contextMenu']>['extraCanvasCommands'];
	hiddenCanvasCommands: ReadonlySet<CanvasContextMenuCommandId>;
	hiddenEditPointsCommands: ReadonlySet<EditPointsCommandId>;
	hiddenDrawingTools: ReadonlySet<ViewerDrawingToolId>;
	elementMenuEnabled: boolean;
	canvasMenuEnabled: boolean;
	hiddenPanels: ReadonlySet<ViewerPanelId>;
	disabledFeatures: ReadonlySet<ViewerFeatureId>;
	hiddenDialogs: ReadonlySet<ViewerDialogId>;
	keyboard: ResolvedKeyboardCustomization;
	/** The host's capability switches as given (`edit` is read by `isEditingAllowed`). */
	capabilities: ViewerCapabilities;
	/** Replaced regions; a hidden region is left out even when the host replaced it. */
	regions: ReadonlyMap<ViewerRegionId, ViewerRegionRenderer>;
	/** Host File tab pages, in the order given. */
	hostBackstagePages: readonly HostBackstagePage[];
}

function addAll<T>(target: Set<T>, values: readonly T[] | undefined): void {
	for (const value of values ?? []) {
		target.add(value);
	}
}

function resolveDialogs(customization: ViewerCustomization): Set<ViewerDialogId> {
	const dialogs = new Set<ViewerDialogId>(customization.hiddenDialogs ?? []);
	const features = new Set(customization.disabledFeatures ?? []);
	if (features.has('collaboration')) {
		dialogs.add('share');
		dialogs.add('broadcast');
	}
	const formats = new Set(customization.hiddenExportFormats ?? []);
	if (VIEWER_EXPORT_FORMAT_IDS.every((id) => formats.has(id))) {
		dialogs.add('export');
	}
	return dialogs;
}

/** The resolved sets that entry points add to. */
interface EntryTargets {
	actions: Set<ToolbarActionId>;
	controls: Set<RibbonControlId>;
	groups: Set<RibbonGroupId>;
	elementCommands: Set<ContextMenuCommandId>;
	shortcuts: Set<EditorKeyActionName>;
	pages: Set<BackstagePage>;
	cards: Set<BackstageCardId>;
}

function hideEntryPoints(targets: EntryTargets, entry: EntryPoints | undefined): void {
	if (!entry) {
		return;
	}
	addAll(targets.actions, entry.actions);
	addAll(targets.controls, entry.controls);
	addAll(targets.groups, entry.groups);
	addAll(targets.elementCommands, entry.elementCommands);
	addAll(targets.shortcuts, entry.shortcuts);
	addAll(targets.pages, entry.pages);
	addAll(targets.cards, entry.cards);
}

/** Normalise `customization` (undefined means "customise nothing"). */
export function resolveCustomization(
	customization: ViewerCustomization | undefined,
): ResolvedCustomization {
	const c = expandCapabilities(customization ?? {});
	const ribbon = resolveRibbonCustomization(c.ribbon);
	const hiddenActions = ribbon.actions;
	const pages = new Set<BackstagePage>(c.backstage?.hiddenPages ?? []);
	const cards = new Set<BackstageCardId>(c.backstage?.hiddenCards ?? []);
	addAll(cards, c.hiddenExportFormats);
	const dialogs = resolveDialogs(c);
	const features = new Set<ViewerFeatureId>(c.disabledFeatures ?? []);
	const panels = new Set<ViewerPanelId>(c.hiddenPanels ?? []);
	const elementCommands = new Set<ContextMenuCommandId>(c.contextMenu?.hiddenElementCommands);
	const keyboard = resolveKeyboardCustomization(c.keyboard);
	const shortcuts = new Set(keyboard.disabled);
	const targets: EntryTargets = {
		actions: hiddenActions,
		controls: ribbon.controls,
		groups: ribbon.groups,
		elementCommands,
		shortcuts,
		pages,
		cards,
	};
	for (const dialog of dialogs) {
		hideEntryPoints(targets, DIALOG_ENTRY_POINTS[dialog]);
	}
	for (const panel of panels) {
		hideEntryPoints(targets, PANEL_ENTRY_POINTS[panel]);
	}
	for (const feature of features) {
		hideEntryPoints(targets, FEATURE_ENTRY_POINTS[feature]);
	}
	const optionsPages = new Set<OptionsPageId>(c.options?.hiddenPages ?? []);
	if (features.has('ai')) {
		optionsPages.add('ai');
	}
	return {
		hiddenActions,
		hiddenInspectorCards: new Set(c.inspector?.hiddenCards ?? []),
		hiddenContextualTabs: ribbon.contextualTabs,
		hiddenRibbonGroups: ribbon.groups,
		hiddenRibbonControls: ribbon.controls,
		hiddenOptionsPages: optionsPages,
		hiddenOptionsSections: new Set(c.options?.hiddenSections ?? []),
		hiddenSettings: new Set(c.options?.hiddenSettings ?? []),
		lockedSettings: { ...c.options?.locked },
		defaultSettings: { ...c.options?.defaults },
		hiddenBackstagePages: pages,
		hiddenBackstageCards: cards,
		hiddenElementCommands: elementCommands,
		extraElementCommands: c.contextMenu?.extraElementCommands ?? [],
		extraCanvasCommands: c.contextMenu?.extraCanvasCommands ?? [],
		hiddenCanvasCommands: new Set(c.contextMenu?.hiddenCanvasCommands ?? []),
		hiddenEditPointsCommands: new Set(c.contextMenu?.hiddenEditPointsCommands ?? []),
		hiddenDrawingTools: new Set(c.hiddenDrawingTools ?? []),
		elementMenuEnabled: c.contextMenu?.disableElementMenu !== true,
		canvasMenuEnabled: c.contextMenu?.disableCanvasMenu !== true,
		hiddenPanels: panels,
		disabledFeatures: features,
		hiddenDialogs: dialogs,
		keyboard: { ...keyboard, disabled: shortcuts },
		capabilities: { ...c.capabilities },
		regions: resolveRegions(c, panels, dialogs, pages),
		hostBackstagePages: c.backstage?.extraPages ?? [],
	};
}

/** The resolution of "no customisation", shared so it is allocated once. */
export const EMPTY_RESOLVED_CUSTOMIZATION: ResolvedCustomization = resolveCustomization(undefined);

/**
 * The `hiddenActions` list a binding threads through its existing ribbon and
 * toolbar gates: the legacy `hiddenActions` prop unioned with the ribbon,
 * dialog and feature rules above. Returns `undefined` when nothing is hidden,
 * which is what those gates treat as "show everything".
 */
export function resolveEffectiveHiddenActions(
	resolved: ResolvedCustomization,
	legacyHiddenActions?: readonly ToolbarActionId[],
): ToolbarActionId[] | undefined {
	if (resolved.hiddenActions.size === 0 && !legacyHiddenActions?.length) {
		return legacyHiddenActions ? [...legacyHiddenActions] : undefined;
	}
	const merged = new Set<ToolbarActionId>(legacyHiddenActions ?? []);
	addAll(merged, [...resolved.hiddenActions]);
	return [...merged];
}

/**
 * True unless the host made the deck read-only (`capabilities.edit: false`).
 * A binding's effective edit state is its own edit prop AND this.
 */
export function isEditingAllowed(resolved: ResolvedCustomization): boolean {
	return resolved.capabilities.edit !== false;
}

/** The host renderer replacing `region`, or undefined to render the built-in. */
export function getRegionRenderer(
	resolved: ResolvedCustomization,
	region: ViewerRegionId,
): ViewerRegionRenderer | undefined {
	return resolved.regions.get(region);
}

/** True when the host replaced `region` (and did not also hide it). */
export function isRegionReplaced(resolved: ResolvedCustomization, region: ViewerRegionId): boolean {
	return resolved.regions.has(region);
}

/** True unless the host removed this chrome region. */
export function isPanelVisible(resolved: ResolvedCustomization, panel: ViewerPanelId): boolean {
	return !resolved.hiddenPanels.has(panel);
}

/** True unless the host switched this feature area off. */
export function isFeatureEnabled(
	resolved: ResolvedCustomization,
	feature: ViewerFeatureId,
): boolean {
	return !resolved.disabledFeatures.has(feature);
}

/** True unless the host removed this dialog (directly, or via a feature). */
export function isDialogAvailable(
	resolved: ResolvedCustomization,
	dialog: ViewerDialogId,
): boolean {
	return !resolved.hiddenDialogs.has(dialog);
}

/**
 * True unless the host switched Edit Points off (the feature, or just its
 * element context-menu entry, which is its only entry point).
 */
export function isEditPointsEnabled(resolved: ResolvedCustomization): boolean {
	return (
		!resolved.disabledFeatures.has('editPoints') &&
		!resolved.hiddenElementCommands.has('edit-points')
	);
}

/** True unless the host removed this Insert > Shapes drawing tool. */
export function isDrawingToolVisible(
	resolved: ResolvedCustomization,
	tool: ViewerDrawingToolId,
): boolean {
	return !resolved.hiddenDrawingTools.has(tool);
}
