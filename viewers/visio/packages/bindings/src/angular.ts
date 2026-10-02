import {
	Component,
	ElementRef,
	EventEmitter,
	Input,
	Output,
	inject,
	type AfterViewInit,
	type OnChanges,
	type OnDestroy,
} from '@angular/core';
import {
	mountFrameworkViewer,
	viewerHandle,
	withEventEmitter,
	type MountedViewer,
	type ViewerCallbacks,
	type ViewerEvents,
	type ViewerProperties,
	type VsdxSource,
	type SvgExportOptions,
	type CurrentPagePrintSnapshotOptions,
} from './common.js';
/** Angular input/output and lifecycle wiring over the one shared browser binding. */
@Component({ selector: 'visio-viewer-host', standalone: true, template: '' })
export class VisioViewerComponent implements AfterViewInit, OnChanges, OnDestroy {
	@Input() document?: ViewerProperties['document'];
	@Input() pageIndex?: number;
	@Input() zoom?: number;
	@Input() showToolbar?: boolean;
	@Input() events?: ViewerCallbacks;
	@Output() documentLoad = new EventEmitter<ViewerEvents['document-load']>();
	@Output() documentError = new EventEmitter<ViewerEvents['document-error']>();
	@Output() pageChange = new EventEmitter<ViewerEvents['page-change']>();
	@Output() zoomChange = new EventEmitter<ViewerEvents['zoom-change']>();
	@Output() shapeSelect = new EventEmitter<ViewerEvents['shape-select']>();
	private readonly host = inject<ElementRef<HTMLElement>>(ElementRef);
	private binding: MountedViewer | undefined;
	private readonly handle = viewerHandle(() => this.binding);
	private readonly outputs = {
		'document-load': this.documentLoad,
		'document-error': this.documentError,
		'page-change': this.pageChange,
		'zoom-change': this.zoomChange,
		'shape-select': this.shapeSelect,
	} satisfies { [K in keyof ViewerEvents]: EventEmitter<ViewerEvents[K]> };
	private options() {
		return withEventEmitter(this, (name, value) => {
			(this.outputs[name] as EventEmitter<unknown>).emit(value);
		});
	}
	ngAfterViewInit() {
		this.binding = mountFrameworkViewer(this.host.nativeElement, this.options());
	}
	ngOnChanges() {
		this.binding?.update(this.options());
	}
	ngOnDestroy() {
		const mounted = this.binding;
		this.binding = undefined;
		mounted?.destroy();
	}
	get element() {
		return this.handle.element;
	}
	get controller() {
		return this.handle.controller;
	}
	load(source: VsdxSource) {
		return this.handle.load(source);
	}
	fit() {
		this.handle.fit();
	}
	setLayerVisibility(pageId: string, layerId: string, visible: boolean | null) {
		this.handle.setLayerVisibility(pageId, layerId, visible);
	}
	resetLayerVisibility(pageId?: string) {
		this.handle.resetLayerVisibility(pageId);
	}
	exportSvg(options?: SvgExportOptions) {
		return this.handle.exportSvg(options);
	}
	createPrintSnapshot(options?: CurrentPagePrintSnapshotOptions) {
		return this.handle.createPrintSnapshot(options);
	}
}
export type { ViewerHandle, ViewerCallbacks, ViewerOptions, ViewerEvents } from './common.js';

// A new shared property must also become an actual Angular input on this class.
const completeInputs: Record<
	Exclude<keyof ViewerProperties, keyof VisioViewerComponent>,
	never
> = {};
void completeInputs;
