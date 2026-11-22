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

/**
 * Excel's native ribbon tabs, as `customization.ribbon.hiddenTabs` names them: File (the tab and
 * its backstage), the main tabs in order, then the contextual tabs.
 */
export const EXCEL_RIBBON_TAB_IDS = [
	'file',
	'home',
	'insert',
	'page-layout',
	'formulas',
	'data',
	'review',
	'view',
	'help',
	'table-design',
	'chart-design',
	'graphics-format',
	'embedded-object',
	'smartart-design',
	'slicer',
	'timeline',
] as const;

/** One native Excel ribbon tab. */
export type ExcelRibbonTabId = (typeof EXCEL_RIBBON_TAB_IDS)[number];

/**
 * Commands that only open a File tab page (Info, Save As without a name, Options): a hidden File
 * tab hides and refuses them, so nothing reaches the backstage it removed.
 */
export const EXCEL_BACKSTAGE_COMMANDS: readonly string[] = [
	'file.info',
	'file.save-as',
	'file.options',
];

/** What a host region (a File tab page) of `<xlsx-editor>` receives. */
export interface ExcelRegionContext extends HostRegionContextBase {
	/** The `<xlsx-editor>` element: its whole public API. */
	readonly editor: HTMLElement;
	readonly fileName: string;
	/** The zero-based index of the active sheet. */
	readonly activeSheet: number;
}

/** Which menu a host command was offered in. */
export type ExcelMenuKind = 'cell' | 'row' | 'column' | 'sheetTab';

/** What a host context-menu command of `<xlsx-editor>` receives: the menu and its selection. */
export interface ExcelMenuContext {
	/** The cell area, a row or column header, or a sheet tab. */
	readonly menu: ExcelMenuKind;
	/** The zero-based index of the sheet. */
	readonly sheet: number;
	readonly sheetName: string;
	/**
	 * The selected range in A1 notation (`A1:B4`; several ranges are comma separated). On the
	 * sheet-tab menu, the selection of the sheet the menu opened on once it is active.
	 */
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
	ExcelMenuContext,
	ExcelRibbonTabId
>;
