/**
 * The cross-cutting rules that turn one host switch into every UI path it
 * implies, kept apart from `customization-resolve.ts` so each table reads as
 * data:
 *
 *  - `capabilities` rewrite into the dialogs, features and command lists they
 *    stand for (`export: false` hides the Export dialog, and so on);
 *  - each dialog and panel lists the toolbar buttons, ribbon controls, menu
 *    entries, shortcuts, File pages and cards that open it, so hiding the
 *    dialog or panel hides its entry points too.
 *
 * @module render/customization/customization-entry-points
 */
import type { BackstagePage } from '../backstage';
import type { BackstageCardId } from '../backstage-cards';
import type { ContextMenuCommandId } from '../context-menu-commands';
import type { EditorKeyActionName } from '../editor-keymap';
import type { ToolbarActionId } from '../toolbar-actions';
import type {
	ViewerCustomization,
	ViewerDialogId,
	ViewerFeatureId,
	ViewerPanelId,
} from './customization-types';
import type { RibbonControlId, RibbonGroupId } from './ribbon-control-ids';

/** Every UI path that opens one dialog or panel. All lists optional. */
export interface EntryPoints {
	actions?: readonly ToolbarActionId[];
	controls?: readonly RibbonControlId[];
	groups?: readonly RibbonGroupId[];
	elementCommands?: readonly ContextMenuCommandId[];
	shortcuts?: readonly EditorKeyActionName[];
	pages?: readonly BackstagePage[];
	cards?: readonly BackstageCardId[];
}

/** Which entry points each dialog owns. */
export const DIALOG_ENTRY_POINTS: Readonly<Record<ViewerDialogId, EntryPoints>> = {
	options: { pages: ['options'], controls: ['help.help.options'] },
	share: { actions: ['share'], pages: ['share'], cards: ['share'] },
	broadcast: { actions: ['broadcast'] },
	print: { pages: ['print'], cards: ['print'] },
	export: {
		actions: ['export'],
		pages: ['export'],
		cards: ['pdf', 'png', 'video', 'gif', 'json', 'copyImage'],
	},
	hyperlink: {
		controls: ['insert.links.link'],
		elementCommands: ['hyperlink'],
		shortcuts: ['hyperlink'],
	},
	equation: { controls: ['insert.symbols.equation'] },
	pasteSpecial: { shortcuts: ['pasteSpecial'] },
	slideTemplates: { controls: ['home.slides.slideTemplates'] },
	passwordProtection: { cards: ['protect'] },
	signatures: { cards: ['signatures'] },
};

/**
 * Which entry points each panel owns. The panels without a row here are
 * always-on chrome with no button that opens them.
 */
export const PANEL_ENTRY_POINTS: Readonly<Partial<Record<ViewerPanelId, EntryPoints>>> = {
	notes: { actions: ['notes'], controls: ['view.show.notes'] },
	selectionPane: { controls: ['view.show.selectionPane'] },
	findReplace: {
		controls: ['home.editing.find', 'home.editing.replace'],
		shortcuts: ['find', 'findReplace'],
	},
	accessibilityPane: {
		controls: ['review.accessibility.check', 'help.help.accessibility'],
	},
	comparePane: { controls: ['review.compare.compare'] },
};

/** Which entry points each switched-off feature removes (beyond its own panels). */
export const FEATURE_ENTRY_POINTS: Readonly<Partial<Record<ViewerFeatureId, EntryPoints>>> = {
	ai: { elementCommands: ['ai-ask', 'ai-fix'] },
	comments: { elementCommands: ['comment'], groups: ['insert.comments', 'review.comments'] },
	presentMode: { actions: ['slideShow'] },
	editPoints: { elementCommands: ['edit-points'] },
};

function union<T>(list: readonly T[] | undefined, extra: readonly T[]): T[] {
	return [...new Set([...(list ?? []), ...extra])];
}

/**
 * Rewrite `capabilities` into the plain lists they stand for, so the rest of
 * the resolver (and every binding) only ever reads hidden / disabled lists.
 * `edit` is not rewritten: it is read directly (`isEditingAllowed`).
 */
export function expandCapabilities(c: ViewerCustomization): ViewerCustomization {
	const caps = c.capabilities;
	if (!caps) {
		return c;
	}
	const dialogs: ViewerDialogId[] = [];
	const features: ViewerFeatureId[] = [];
	if (caps.export === false) {
		dialogs.push('export');
	}
	if (caps.print === false) {
		dialogs.push('print');
	}
	if (caps.share === false) {
		dialogs.push('share', 'broadcast');
	}
	if (caps.present === false) {
		features.push('presentMode');
	}
	if (caps.comment === false) {
		features.push('comments');
	}
	const next: ViewerCustomization = {
		...c,
		hiddenDialogs: union(c.hiddenDialogs, dialogs),
		disabledFeatures: union(c.disabledFeatures, features),
	};
	if (caps.copy === false) {
		next.contextMenu = {
			...c.contextMenu,
			hiddenElementCommands: union(c.contextMenu?.hiddenElementCommands, [
				'copy',
				'cut',
				'save-as-picture',
			]),
		};
		next.keyboard = { ...c.keyboard, disabled: union(c.keyboard?.disabled, ['copy', 'cut']) };
		next.ribbon = {
			...c.ribbon,
			hiddenButtons: union(c.ribbon?.hiddenButtons, [
				'home.clipboard.copy',
				'home.clipboard.cut',
			] as const),
		};
		next.hiddenExportFormats = union(c.hiddenExportFormats, ['copyImage']);
	}
	return next;
}
