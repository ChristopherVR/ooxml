/**
 * `<visio-viewer>`'s `customization` property: the shared contract of `ooxml-core/embedding`
 * bound to Visio. Capabilities hide their ribbon, menu and File tab commands and refuse the
 * matching actions however they are raised; `backstage.extraPages` and
 * `contextMenu.extraCommands` reach the File tab and the shape and page menus.
 */
import {
	capabilityHiddenActions,
	isCapabilityEnabled,
	type CapabilityActionMap,
} from 'ooxml-core/embedding';
import {
	VISIO_CAPABILITY_BACKSTAGE_PAGES,
	VISIO_CAPABILITY_COMMANDS,
	type VisioBackstagePage,
	type VisioCapabilityId,
	type VisioCustomization,
	type VisioMenuCommand,
	type VisioMenuContext,
} from 'ooxml-core/visio/ui';
import { syncHostMenuChildren, tidyMenuSeparators } from '../embedding/host-menu';
import type { ViewerController } from './controller';
import type { VisioRibbonAction } from './ribbon-action';

/** A customisation resolved into what each render site asks. */
export interface ResolvedVisioCustomization {
	readonly canEdit: boolean;
	/** File tab entries to remove. */
	readonly backstagePages: ReadonlySet<string>;
	/** Ribbon, menu and File tab commands (`command` / `data-backstage-action` ids) to hide. */
	readonly commands: ReadonlySet<string>;
	readonly disabled: ReadonlySet<VisioCapabilityId>;
	readonly extraPages: readonly VisioBackstagePage[];
	readonly menuCommands: readonly VisioMenuCommand[];
}

/** Resolves a host customisation (none: everything on, nothing added). */
export function resolveVisioCustomization(
	customization?: VisioCustomization | null,
): ResolvedVisioCustomization {
	const capabilities = customization?.capabilities;
	const set = (map: CapabilityActionMap<VisioCapabilityId, string>) =>
		new Set(capabilityHiddenActions(capabilities, map));
	const ids = Object.keys(capabilities ?? {}) as VisioCapabilityId[];
	return {
		canEdit: isCapabilityEnabled(capabilities, 'edit'),
		backstagePages: set(VISIO_CAPABILITY_BACKSTAGE_PAGES),
		commands: set(VISIO_CAPABILITY_COMMANDS),
		disabled: new Set(ids.filter((id) => !isCapabilityEnabled(capabilities, id))),
		extraPages: customization?.backstage?.extraPages ?? [],
		menuCommands: customization?.contextMenu?.extraCommands ?? [],
	};
}

export const EMPTY_VISIO_CUSTOMIZATION: ResolvedVisioCustomization = resolveVisioCustomization();

/** The capability an action needs, when a capability governs it. */
export function visioActionCapability(action: VisioRibbonAction): VisioCapabilityId | undefined {
	switch (action.type) {
		case 'clipboard':
			return action.operation === 'paste' ? undefined : 'copy';
		case 'search':
		case 'replace':
			return 'find';
		case 'review':
			return action.command === 'new-comment' || action.command === 'page-comment'
				? 'comment'
				: undefined;
		default:
			return undefined;
	}
}

/** True when a disabled capability refuses `action`. */
export function visioActionBlocked(
	resolved: ResolvedVisioCustomization,
	action: VisioRibbonAction,
): boolean {
	const capability = visioActionCapability(action);
	return capability !== undefined && resolved.disabled.has(capability);
}

/** Every id a capability can hide, so a removed capability shows its commands again. */
const GOVERNED = new Set(Object.values(VISIO_CAPABILITY_COMMANDS).flat());

const HIDDEN_BY_CAPABILITY = 'data-capability-hidden';

/** Hides (or shows again) the ribbon, menu and File tab commands the capabilities govern. */
export function applyVisioCommandVisibility(
	root: ParentNode,
	resolved: ResolvedVisioCustomization,
): void {
	const selector = '[command], [data-menu], [data-backstage-action]';
	for (const element of root.querySelectorAll<HTMLElement>(selector)) {
		const id =
			element.getAttribute('command') ??
			element.dataset.menu ??
			element.dataset.backstageAction ??
			'';
		if (!GOVERNED.has(id)) continue;
		// Only undo what a capability hid, so a command the viewer hides itself stays hidden.
		if (resolved.commands.has(id)) {
			element.setAttribute(HIDDEN_BY_CAPABILITY, '');
			element.hidden = true;
		} else if (element.hasAttribute(HIDDEN_BY_CAPABILITY)) {
			element.removeAttribute(HIDDEN_BY_CAPABILITY);
			element.hidden = false;
		}
	}
}

/** The element's customisation: the host object, its resolution and the requested `canEdit`. */
export class ViewerCustomizationState {
	input: VisioCustomization | null = null;
	resolved: ResolvedVisioCustomization = EMPTY_VISIO_CUSTOMIZATION;
	/** The `canEdit` property as the host set it. */
	requestedCanEdit = true;
	/** Effective: the property, unless `capabilities.edit` is off. */
	get canEdit(): boolean {
		return this.requestedCanEdit && this.resolved.canEdit;
	}
	/** Replaces the customisation; false when it is the same object. */
	set(value: VisioCustomization | null): boolean {
		if (value === this.input) return false;
		this.input = value;
		this.resolved = resolveVisioCustomization(value);
		return true;
	}
}

/**
 * Adds the host commands to the shape or page menu just before it opens, and tidies the
 * separators that capability-hidden entries left.
 */
export function syncVisioHostMenu(
	menu: HTMLElement,
	target: 'shape' | 'page',
	resolved: ResolvedVisioCustomization,
	controller: Pick<ViewerController, 'state'>,
): void {
	const { state } = controller;
	const pageId = state.document?.pages[state.pageIndex]?.id;
	const context: VisioMenuContext = {
		target,
		pageIndex: state.pageIndex,
		...(pageId ? { pageId } : {}),
		shapes: target === 'shape' ? [...state.selectedShapes] : [],
	};
	syncHostMenuChildren(menu, resolved.menuCommands, context);
	tidyMenuSeparators(menu);
}

/**
 * The Quick Access Toolbar commands to show: the File > Options choices, less the ones a
 * disabled capability removes (Save with `save`), as Word drops its Quick Access Save.
 */
export function quickAccessShown<T extends object>(
	behaviour: T,
	resolved: ResolvedVisioCustomization,
): T & { 'qat-save'?: boolean } {
	return resolved.commands.has('download') ? { ...behaviour, 'qat-save': false } : behaviour;
}
