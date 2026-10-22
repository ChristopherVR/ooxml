import {
	Component,
	DestroyRef,
	ElementRef,
	EventEmitter,
	Injectable,
	Input,
	Output,
	computed,
	effect,
	untracked,
	inject,
	signal,
	type AfterViewInit,
	type OnChanges,
	type OnDestroy,
	type OnInit,
} from '@angular/core';
import {
	buildingBlock,
	prepareDiagramBlock,
	createBuildingBlockSession,
	effectiveMode,
	type BuildingBlockOptions,
} from '../../bindings/src/building-blocks';
import {
	toolbarGroups,
	type ToolbarButton,
	type ToolbarControl,
	type ToolbarGroupId,
	type ToolbarSearch,
	type ToolbarText,
} from '../../bindings/src/toolbar-controls';
@Injectable()
export class ViewerBuildingBlocks {
	private readonly session = createBuildingBlockSession();
	readonly controller = this.session.controller;
	private readonly current = signal(this.controller.state);
	readonly state = this.current.asReadonly();
	private readonly latest = signal<BuildingBlockOptions>({});
	readonly toolbarProps = computed(() => {
		this.state();
		return this.session.toolbar(this.latest());
	});
	readonly loading = computed(() => this.state().loading);
	readonly error = computed(() => this.state().error ?? this.state().edit.error);
	/** Effective permission: host `canEdit` and a drawing that can be saved back. */
	readonly mode = computed(() => effectiveMode(this.state()));
	readonly documentChange = new EventEmitter<NonNullable<typeof this.controller.state.document>>();
	readonly dirtyChange = new EventEmitter<boolean>();
	readonly documentError = new EventEmitter<Error>();
	readonly canvasProps = { mount: (node: HTMLElement | null) => this.session.mount(node) };
	readonly inspectorProps = { controller: this.controller };
	readonly shapesProps = computed(() => ({
		controller: this.controller,
		locale: this.latest().locale,
	}));
	constructor() {
		this.update({});
		const stop = this.controller.subscribe((state) => this.current.set(state));
		inject(DestroyRef).onDestroy(() => {
			stop();
			this.session.destroy();
		});
	}
	/** Acknowledge a host save: clears dirty and emits `dirtyChange(false)`. */
	markClean() {
		this.session.markClean();
	}
	update(options: BuildingBlockOptions) {
		this.latest.set(options);
		this.session.update({
			...options,
			onDocumentChange: (document) => {
				this.documentChange.emit(document);
				options.onDocumentChange?.(document);
			},
			onDirtyChange: (dirty) => {
				this.dirtyChange.emit(dirty);
				options.onDirtyChange?.(dirty);
			},
			onDocumentError: (error) => {
				this.documentError.emit(error);
				options.onDocumentError?.(error);
			},
		});
	}
}
/** Connect host signals to the scoped owner within an Angular injection context. */
export function createViewerBuildingBlocks(options: () => BuildingBlockOptions = () => ({})) {
	const blocks = inject(ViewerBuildingBlocks);
	effect(() => {
		const next = options();
		untracked(() => blocks.update(next));
	});
	return blocks;
}

@Component({
	selector: 'visio-building-blocks',
	standalone: true,
	template: '<ng-content />',
	providers: [ViewerBuildingBlocks],
})
export class ViewerBuildingBlocksComponent implements OnChanges, OnInit {
	readonly blocks = inject(ViewerBuildingBlocks);
	@Input() content?: BuildingBlockOptions['content'];
	/** Host edit permission. Read-only until `true`. */
	@Input() canEdit = false;
	@Input() locale?: string;
	@Input() onOpenFile?: () => void;
	@Input() onOpenPageSetup?: () => void;
	@Input() onOpenSettings?: () => void;
	@Input() onOpenShareDialog?: () => void;
	@Output() readonly documentChange = this.blocks.documentChange;
	@Output() readonly dirtyChange = this.blocks.dirtyChange;
	@Output() readonly documentError = this.blocks.documentError;
	ngOnInit() {
		this.ngOnChanges();
	}
	ngOnChanges() {
		this.blocks.update({
			content: this.content,
			canEdit: this.canEdit,
			locale: this.locale,
			onOpenFile: this.onOpenFile,
			onOpenPageSetup: this.onOpenPageSetup,
			onOpenSettings: this.onOpenSettings,
			onOpenShareDialog: this.onOpenShareDialog,
		});
	}
}
@Component({ selector: 'visio-diagram-canvas', standalone: true, template: '' })
export class DiagramCanvasComponent implements AfterViewInit, OnDestroy {
	readonly blocks = inject(ViewerBuildingBlocks);
	private readonly host = inject<ElementRef<HTMLElement>>(ElementRef);
	ngAfterViewInit() {
		this.blocks.canvasProps.mount(this.host.nativeElement);
	}
	ngOnDestroy() {
		this.blocks.canvasProps.mount(null);
	}
}
@Component({
	selector: 'visio-diagram-toolbar',
	standalone: true,
	template: `<div role="toolbar">
		@for (group of groups(); track group.id) {
			<span [attr.data-group]="group.id">
				@for (control of group.controls; track control.id) {
					@switch (control.kind) {
						@case ('text') {
							<span [attr.data-control]="control.id">{{ asText(control).text }}</span>
						}
						@case ('button') {
							<button
								type="button"
								[attr.data-control]="control.id"
								[disabled]="asButton(control).disabled"
								[attr.aria-pressed]="asButton(control).pressed ?? null"
								(click)="asButton(control).run()"
							>
								{{ asButton(control).label }}
							</button>
						}
						@default {
							<span data-control="search">
								<input
									type="search"
									[attr.aria-label]="asSearch(control).label"
									[attr.placeholder]="asSearch(control).label"
									[value]="asSearch(control).query"
									(input)="asSearch(control).change(value($event))"
								/>
								<span aria-live="polite">{{ asSearch(control).summary }}</span>
								<button
									type="button"
									[attr.data-control]="asSearch(control).previous.id"
									[disabled]="asSearch(control).previous.disabled"
									(click)="asSearch(control).previous.run()"
								>
									{{ asSearch(control).previous.label }}
								</button>
								<button
									type="button"
									[attr.data-control]="asSearch(control).next.id"
									[disabled]="asSearch(control).next.disabled"
									(click)="asSearch(control).next.run()"
								>
									{{ asSearch(control).next.label }}
								</button>
							</span>
						}
					}
				}
			</span>
		}
	</div>`,
})
export class ToolbarComponent implements AfterViewInit {
	private readonly host = inject<ElementRef<HTMLElement>>(ElementRef);
	/** Groups to show; omit for all. Hosts may build their own controls from toolbarProps. */
	@Input() include?: readonly ToolbarGroupId[];
	ngAfterViewInit() {
		prepareDiagramBlock(this.host.nativeElement.querySelector('div'), 'toolbar');
	}
	readonly blocks = inject(ViewerBuildingBlocks);
	groups() {
		return toolbarGroups(this.blocks.toolbarProps(), this.include);
	}
	asText(control: ToolbarControl) {
		return control as ToolbarText;
	}
	asButton(control: ToolbarControl) {
		return control as ToolbarButton;
	}
	asSearch(control: ToolbarControl) {
		return control as ToolbarSearch;
	}
	value(event: Event) {
		return (event.target as HTMLInputElement).value;
	}
}
@Component({
	selector: 'visio-page-navigation',
	standalone: true,
	template: `<select
		[attr.aria-label]="blocks.toolbarProps().labels['Page']"
		[value]="blocks.toolbarProps().pageIndex"
		(change)="change($event)"
	>
		@for (page of blocks.toolbarProps().pages; track page.id; let index = $index) {
			<option [value]="index">{{ page.name }}</option>
		}
	</select>`,
})
export class PageNavigationComponent implements AfterViewInit {
	private readonly host = inject<ElementRef<HTMLElement>>(ElementRef);
	ngAfterViewInit() {
		prepareDiagramBlock(this.host.nativeElement.querySelector('select'), 'tabs');
	}
	readonly blocks = inject(ViewerBuildingBlocks);
	change(event: Event) {
		this.blocks.toolbarProps().onPageChange(Number((event.target as HTMLSelectElement).value));
	}
}
@Component({ selector: 'visio-diagram-inspector', standalone: true, template: '' })
export class InspectorComponent implements AfterViewInit, OnDestroy {
	readonly blocks = inject(ViewerBuildingBlocks);
	private readonly host = inject<ElementRef<HTMLElement>>(ElementRef);
	private stop: (() => void) | undefined;
	ngAfterViewInit() {
		this.stop = buildingBlock('mountDiagramInspector')(
			this.host.nativeElement,
			this.blocks.controller,
		);
	}
	ngOnDestroy() {
		this.stop?.();
	}
}
@Component({ selector: 'visio-shapes-pane', standalone: true, template: '' })
export class ShapesPaneComponent implements AfterViewInit, OnDestroy {
	readonly blocks = inject(ViewerBuildingBlocks);
	private readonly host = inject<ElementRef<HTMLElement>>(ElementRef);
	private stop: (() => void) | undefined;
	ngAfterViewInit() {
		this.stop = buildingBlock('mountDiagramShapes')(
			this.host.nativeElement,
			this.blocks.controller,
			{ locale: this.blocks.shapesProps().locale },
		);
	}
	ngOnDestroy() {
		this.stop?.();
	}
}
export type {
	BuildingBlockOptions,
	BuildingBlockHostCallbacks,
	DiagramToolbarProps,
} from '../../bindings/src/building-blocks';
export type { ToolbarGroupId } from '../../bindings/src/toolbar-controls';
