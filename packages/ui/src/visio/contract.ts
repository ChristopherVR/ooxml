import type { VisioDocument } from 'ooxml-core/visio';

export interface ViewerProperties {
	document: VisioDocument | null;
	pageIndex: number;
	zoom: number;
	showToolbar: boolean;
}
export interface ViewerEvents {
	'document-load': VisioDocument;
	'document-change': {
		document: VisioDocument;
		dirty: boolean;
		kind: 'edit' | 'undo' | 'redo' | 'remote';
	};
	'document-error': Error;
	'page-change': number;
	'zoom-change': number;
	'shape-select': { id: string; name: string; pageId?: string } | null;
}
export type ViewerCallbacks = {
	[K in keyof ViewerEvents]?: (detail: ViewerEvents[K]) => void;
};
export const propertyKeys = [
	'document',
	'pageIndex',
	'zoom',
	'showToolbar',
] as const satisfies readonly (keyof ViewerProperties)[];
export const eventKeys = [
	'document-load',
	'document-change',
	'document-error',
	'page-change',
	'zoom-change',
	'shape-select',
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
