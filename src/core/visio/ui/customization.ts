/**
 * Visio's half of the shared customization contract (`ooxml-core/embedding`): the capabilities
 * `<visio-viewer>` honours, what each one removes, and the `customization` property's type every
 * binding forwards. The renderer of a host page receives an `HTMLElement` container, as in the
 * other editors; nothing here touches the DOM.
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
import type { VisioShapeSelection } from './contract';

/** The capabilities Visio honours. */
export type VisioCapabilityId = Extract<
	EditorCapabilityId,
	'edit' | 'export' | 'print' | 'share' | 'copy' | 'comment' | 'find' | 'open' | 'save' | 'settings'
>;

/** Every Visio capability, for hosts that build a policy screen. */
export const VISIO_CAPABILITY_IDS = [
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
] as const satisfies readonly VisioCapabilityId[];

/** Per-capability switches of `<visio-viewer>`. */
export type VisioCapabilities = EditorCapabilities<VisioCapabilityId>;

/** The File tab entries each disabled capability removes. */
export const VISIO_CAPABILITY_BACKSTAGE_PAGES: CapabilityActionMap<VisioCapabilityId, string> = {
	open: ['new', 'open'],
	save: ['save', 'save-as'],
	export: ['export'],
	print: ['print'],
	share: ['share'],
	settings: ['options'],
};

/**
 * The ribbon, context-menu and File tab commands (their `command`, menu or File tab action ids) each disabled capability
 * hides. The editor also refuses the matching actions however they are raised (a shortcut, a
 * menu, the API).
 */
export const VISIO_CAPABILITY_COMMANDS: CapabilityActionMap<VisioCapabilityId, string> = {
	copy: ['cut', 'copy', 'ctx-cut', 'ctx-copy'],
	comment: ['new-comment', 'ctx-comment', 'ctx-page-comment'],
	find: ['find', 'find-item', 'replace'],
	open: ['open', 'new-blank'],
	save: ['download'],
	export: ['export-svg', 'export-png', 'export-pdf'],
	print: ['print'],
	share: ['email'],
};

/** What a host region (a File tab page) of `<visio-viewer>` receives. */
export interface VisioRegionContext extends HostRegionContextBase {
	/** The `<visio-viewer>` element: its whole public API. */
	readonly viewer: HTMLElement;
	/** The zero-based index of the page on screen. */
	readonly pageIndex: number;
	readonly pageCount: number;
}

/** What a host context-menu command of `<visio-viewer>` receives. */
export interface VisioMenuContext {
	/** `shape` when the menu opened on a shape, `page` on empty paper. */
	readonly target: 'shape' | 'page';
	readonly pageIndex: number;
	/** The page's id, when a drawing is open. */
	readonly pageId?: string;
	/** The selection the menu acts on (empty on the page menu). */
	readonly shapes: readonly VisioShapeSelection[];
}

export type VisioRegionRenderer = HostRegionRenderer<HTMLElement, VisioRegionContext>;
export type VisioBackstagePage = HostBackstagePage<VisioRegionRenderer>;
export type VisioMenuCommand = HostMenuCommand<VisioMenuContext>;

/** The whole `customization` object of `<visio-viewer>`. Every field is optional. */
export type VisioCustomization = EditorCustomization<
	VisioCapabilityId,
	VisioRegionRenderer,
	VisioMenuContext
>;
