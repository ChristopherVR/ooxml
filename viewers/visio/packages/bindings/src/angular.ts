import {
	Component,
	ElementRef,
	EventEmitter,
	Input,
	Output,
	inject,
	signal,
	type AfterViewInit,
	type OnChanges,
	type OnDestroy,
} from '@angular/core';
import { installHandle } from 'ooxml-ui/bindings';
import {
	mountFrameworkViewer,
	VIEWER_CONTRACT,
	VIEWER_HANDLE_OPTIONS,
	viewerHandle,
	viewerStateSource,
	withEventEmitter,
	type MountedViewer,
	type ViewerCallbacks,
	type ViewerEvents,
	type ViewerHandle,
	type ViewerProperties,
	type ViewerState,
} from './common';
/** The handle members, defined on the prototype from `VIEWER_CONTRACT.handle`. */
// oxlint-disable-next-line typescript/no-unsafe-declaration-merging -- members come from installHandle.
export interface VisioViewerComponent extends ViewerHandle {}
/** Angular input/output and lifecycle wiring over the one shared browser binding. */
// oxlint-disable-next-line typescript/no-unsafe-declaration-merging -- see the interface above.
@Component({ selector: 'visio-viewer-host', standalone: true, template: '' })
export class VisioViewerComponent implements AfterViewInit, OnChanges, OnDestroy {
	@Input() document?: ViewerProperties['document'];
	@Input() pageIndex?: number;
	@Input() zoom?: number;
	@Input() showToolbar?: boolean;
	@Input() canEdit?: boolean;
	@Input() readOnly?: boolean;
	@Input() ribbonAddIns?: ViewerProperties['ribbonAddIns'];
	@Input() collaboration?: ViewerProperties['collaboration'];
	@Input() storage?: ViewerProperties['storage'];
	@Input() customization?: ViewerProperties['customization'];
	@Input() events?: ViewerCallbacks;
	@Output() documentLoad = new EventEmitter<ViewerEvents['document-load']>();
	@Output() documentChange = new EventEmitter<ViewerEvents['document-change']>();
	@Output() documentError = new EventEmitter<ViewerEvents['document-error']>();
	@Output() pageChange = new EventEmitter<ViewerEvents['page-change']>();
	@Output() zoomChange = new EventEmitter<ViewerEvents['zoom-change']>();
	@Output() shapeSelect = new EventEmitter<ViewerEvents['shape-select']>();
	@Output() selectionChange = new EventEmitter<ViewerEvents['selection-change']>();
	@Output() ribbonAddIn = new EventEmitter<ViewerEvents['office-ribbon-add-in']>();
	@Output() collaborationChange = new EventEmitter<ViewerEvents['collaboration-change']>();
	/** Notification only: to cancel a File command, return false from `events['file-command']`. */
	@Output() fileCommand = new EventEmitter<ViewerEvents['file-command']>();
	private readonly host = inject<ElementRef<HTMLElement>>(ElementRef);
	private binding: MountedViewer | undefined;
	private readonly handle = viewerHandle(() => this.binding);
	static {
		installHandle(
			this.prototype,
			VIEWER_CONTRACT.handle,
			(self) => self.binding,
			VIEWER_HANDLE_OPTIONS,
		);
	}
	private stopState = () => {};
	private readonly viewerState = signal<ViewerState | null>(null);
	/** Reactive viewer state as an Angular signal (`viewer.state()?.pageIndex`); `null` until mounted. */
	readonly state = this.viewerState.asReadonly();
	private readonly outputs = {
		'document-load': this.documentLoad,
		'document-change': this.documentChange,
		'document-error': this.documentError,
		'page-change': this.pageChange,
		'zoom-change': this.zoomChange,
		'shape-select': this.shapeSelect,
		'selection-change': this.selectionChange,
		'office-ribbon-add-in': this.ribbonAddIn,
		'collaboration-change': this.collaborationChange,
		'file-command': this.fileCommand,
	} satisfies { [K in keyof ViewerEvents]: EventEmitter<ViewerEvents[K]> };
	private options() {
		return withEventEmitter(this, (name, value) => {
			(this.outputs[name] as EventEmitter<unknown>).emit(value);
		});
	}
	ngAfterViewInit() {
		this.binding = mountFrameworkViewer(this.host.nativeElement, this.options());
		const source = viewerStateSource(() => this.handle);
		this.viewerState.set(source.getSnapshot());
		this.stopState = source.subscribe(() => this.viewerState.set(source.getSnapshot()));
	}
	ngOnChanges() {
		this.binding?.update(this.options());
	}
	ngOnDestroy() {
		this.stopState();
		this.stopState = () => {};
		this.viewerState.set(null);
		const mounted = this.binding;
		this.binding = undefined;
		mounted?.destroy();
	}
}
export type {
	ViewerState,
	ViewerHandle,
	ViewerCallbacks,
	ViewerOptions,
	ViewerEvents,
	VisioShapeSelection,
	CreateVsdxOptions,
	ViewerEditState,
	VsdxExportResult,
	RibbonAddInCommand,
	RibbonAddInGroup,
	RibbonAddInTab,
} from './common';

// A new shared property must also become an actual Angular input on this class.
const completeInputs: Record<
	Exclude<keyof ViewerProperties, keyof VisioViewerComponent>,
	never
> = {};
void completeInputs;
