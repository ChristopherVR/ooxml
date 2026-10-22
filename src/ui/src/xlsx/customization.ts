/**
 * `<xlsx-editor>`'s `customization` property: the shared contract of `ooxml-core/embedding`
 * bound to Excel. Capabilities hide their commands through the command registry (so the ribbon,
 * Tell me, the context menus, the shortcuts and the title bar all refuse them) and remove their
 * File tab pages; `backstage.extraPages` and `contextMenu.extraCommands` reach the File tab and
 * the cell menu.
 */
import {
	capabilityHiddenActions,
	isCapabilityEnabled,
	type CapabilityActionMap,
} from 'ooxml-core/embedding';
import {
	EXCEL_CAPABILITY_BACKSTAGE_PAGES,
	EXCEL_CAPABILITY_COMMANDS,
	type EditorContext,
	type ExcelBackstagePage,
	type ExcelCapabilityId,
	type ExcelCustomization,
	type ExcelMenuCommand,
	type ExcelMenuContext,
	type ExcelMenuKind,
	activeRef,
	selectionRef,
} from 'ooxml-core/xlsx/ui';

/** A customisation resolved into what each render site asks. */
export interface ResolvedExcelCustomization {
	readonly canEdit: boolean;
	/** Command ids the capabilities hide, unioned with `hiddenActions` by the registry. */
	readonly commands: ReadonlySet<string>;
	readonly backstagePages: ReadonlySet<string>;
	readonly extraPages: readonly ExcelBackstagePage[];
	readonly menuCommands: readonly ExcelMenuCommand[];
}

/** Resolves a host customisation (none: everything on, nothing added). */
export function resolveExcelCustomization(
	customization?: ExcelCustomization | null,
): ResolvedExcelCustomization {
	const capabilities = customization?.capabilities;
	const set = (map: CapabilityActionMap<ExcelCapabilityId, string>) =>
		new Set(capabilityHiddenActions(capabilities, map));
	return {
		canEdit: isCapabilityEnabled(capabilities, 'edit'),
		commands: set(EXCEL_CAPABILITY_COMMANDS),
		backstagePages: set(EXCEL_CAPABILITY_BACKSTAGE_PAGES),
		extraPages: customization?.backstage?.extraPages ?? [],
		menuCommands: customization?.contextMenu?.extraCommands ?? [],
	};
}

export const EMPTY_EXCEL_CUSTOMIZATION: ResolvedExcelCustomization = resolveExcelCustomization();

const registry = new WeakMap<EditorContext, () => ResolvedExcelCustomization>();

/** Lets the menus of `ctx` read its editor's customisation (the editor core registers it). */
export function provideCustomization(
	ctx: EditorContext,
	read: () => ResolvedExcelCustomization,
): void {
	registry.set(ctx, read);
}

/** The customisation of the editor `ctx` belongs to. */
export function customizationOf(ctx: EditorContext): ResolvedExcelCustomization {
	return registry.get(ctx)?.() ?? EMPTY_EXCEL_CUSTOMIZATION;
}

/**
 * The `host` option of `openContextMenu` for one menu: the editor's host commands with the
 * context of this opening, or nothing when the host added none.
 */
export function hostMenu(
	ctx: EditorContext,
	menu: ExcelMenuKind,
	sheet: number,
): { host?: { commands: readonly ExcelMenuCommand[]; context: ExcelMenuContext } } {
	const commands = customizationOf(ctx).menuCommands;
	if (!commands.length) return {};
	const selection = ctx.selection.get();
	const context: ExcelMenuContext = {
		menu,
		sheet,
		sheetName: ctx.workbook()?.sheets[sheet]?.name ?? '',
		ref: selectionRef(selection),
		active: activeRef(selection),
		readOnly: ctx.readOnly(),
	};
	return { host: { commands, context } };
}
