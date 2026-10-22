/**
 * Excel's half of the shared customization contract (`ooxml-core/embedding`): the capabilities
 * `<xlsx-editor>` honours, what each one removes, and the `customization` property's type every
 * binding forwards. A disabled capability's commands are hidden through the command registry,
 * so the ribbon, the context menus, the shortcuts and the File tab all refuse them.
 */
import type {
	CapabilityActionMap,
	EditorCapabilities,
	EditorCapabilityId,
	EditorCustomization,
	HostBackstagePage,
	HostMenuCommand,
	HostRegionContextBase,
	HostRegionRenderer,
} from '../../embedding/index';

/** The capabilities Excel honours. */
export type ExcelCapabilityId = Extract<
	EditorCapabilityId,
	'edit' | 'export' | 'print' | 'share' | 'copy' | 'comment' | 'find' | 'open' | 'save' | 'settings'
>;

/** Every Excel capability, for hosts that build a policy screen. */
export const EXCEL_CAPABILITY_IDS = [
	'edit',
	'export',
	'print',
	'share',
	'copy',
	'comment',
	'find',
	'open',
	'save',
	'settings',
] as const satisfies readonly ExcelCapabilityId[];

/** Per-capability switches of `<xlsx-editor>`. */
export type ExcelCapabilities = EditorCapabilities<ExcelCapabilityId>;

/** The command ids (`hiddenActions` ids) each disabled capability hides and refuses to run. */
export const EXCEL_CAPABILITY_COMMANDS: CapabilityActionMap<ExcelCapabilityId, string> = {
	open: ['file.new', 'file.open'],
	save: ['file.save'],
	export: ['file.save-as', 'file.export', 'file.export-csv'],
	print: ['file.print'],
	share: ['file.share'],
	settings: ['file.options'],
	copy: ['edit.copy', 'edit.cut', 'home.copy', 'home.cut'],
	comment: ['insert.comment', 'review.new-comment', 'review.delete-comment'],
	find: ['home.find', 'home.replace'],
};

/** The File tab entries each disabled capability removes. */
export const EXCEL_CAPABILITY_BACKSTAGE_PAGES: CapabilityActionMap<ExcelCapabilityId, string> = {
	open: ['new', 'open'],
	save: ['save'],
	export: ['saveAs', 'export'],
	print: ['print'],
	share: ['share'],
	settings: ['options'],
};

/** What a host region (a File tab page) of `<xlsx-editor>` receives. */
export interface ExcelRegionContext extends HostRegionContextBase {
	/** The `<xlsx-editor>` element: its whole public API. */
	readonly editor: HTMLElement;
	readonly fileName: string;
	/** The zero-based index of the active sheet. */
	readonly activeSheet: number;
}

/** What a host context-menu command of `<xlsx-editor>` receives: the cell menu's selection. */
export interface ExcelMenuContext {
	/** The zero-based index of the sheet. */
	readonly sheet: number;
	readonly sheetName: string;
	/** The selected range in A1 notation (`A1:B4`; several ranges are comma separated). */
	readonly ref: string;
	/** The active cell in A1 notation. */
	readonly active: string;
	readonly readOnly: boolean;
}

export type ExcelRegionRenderer = HostRegionRenderer<HTMLElement, ExcelRegionContext>;
export type ExcelBackstagePage = HostBackstagePage<ExcelRegionRenderer>;
export type ExcelMenuCommand = HostMenuCommand<ExcelMenuContext>;

/** The whole `customization` object of `<xlsx-editor>`. Every field is optional. */
export type ExcelCustomization = EditorCustomization<
	ExcelCapabilityId,
	ExcelRegionRenderer,
	ExcelMenuContext
>;
