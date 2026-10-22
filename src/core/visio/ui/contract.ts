import type { RibbonAddInCommandDetail, RibbonAddInTab } from '../../ribbon/index';
import type { CollaborationState } from '../../collab/index';
import type { VisioCollaborationOptions } from './collab-options';
import type { VisioDocument } from '../index';

export interface ViewerProperties {
	canEdit?: boolean;
	document: VisioDocument | null;
	pageIndex: number;
	zoom: number;
	showToolbar: boolean;
	/** Tabs the host adds after Help, as an Office add-in does. */
	ribbonAddIns: readonly RibbonAddInTab[];
	/** Shares the open drawing in a room while set; null leaves it (the shared embedding contract). */
	collaboration: VisioCollaborationOptions | null;
}
/** A page-scoped shape identity. Omitted pageId means the currently displayed page. */
export interface VisioShapeSelection {
	readonly id: string;
	readonly name: string;
	readonly pageId?: string;
}
export interface ViewerEvents {
	'document-load': VisioDocument;
	'document-change': {
		document: VisioDocument;
		dirty: boolean;
		kind: 'edit' | 'undo' | 'redo' | 'remote' | 'save';
	};
	'document-error': Error;
	'page-change': number;
	'zoom-change': number;
	'shape-select': VisioShapeSelection | null;
	/** Ordered immutable selection; the first entry is the primary selectedShape. */
	'selection-change': readonly VisioShapeSelection[];
	/** A command of a host tab (`ribbonAddIns`) was chosen. */
	'office-ribbon-add-in': RibbonAddInCommandDetail;
	/** Sharing started, stopped or its status, sync or people changed. */
	'collaboration-change': CollaborationState;
}
export type ViewerCallbacks = {
	[K in keyof ViewerEvents]?: (detail: ViewerEvents[K]) => void;
};
export const propertyKeys = [
	'canEdit',
	'document',
	'pageIndex',
	'zoom',
	'showToolbar',
	'ribbonAddIns',
	'collaboration',
] as const satisfies readonly (keyof ViewerProperties)[];
export const eventKeys = [
	'document-load',
	'document-change',
	'document-error',
	'page-change',
	'zoom-change',
	'shape-select',
	'selection-change',
	'office-ribbon-add-in',
	'collaboration-change',
] as const satisfies readonly (keyof ViewerEvents)[];
// Adding a contract member must also update the inventories used by every adapter.
const allProperties: Record<
	Exclude<keyof ViewerProperties, (typeof propertyKeys)[number]>,
	never
> = {};
const allEvents: Record<Exclude<keyof ViewerEvents, (typeof eventKeys)[number]>, never> = {};
void allProperties;
void allEvents;
export type VsdxSource = ArrayBuffer | Uint8Array | Blob;
export interface ViewerOptions extends Partial<ViewerProperties> {
	events?: ViewerCallbacks;
}
