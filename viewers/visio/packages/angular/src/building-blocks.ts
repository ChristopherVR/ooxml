import {
	Component,
	DestroyRef,
	ElementRef,
	EventEmitter,
	Injectable,
	Input,
	Output,
	computed,
	inject,
	signal,
	type AfterViewInit,
	type OnChanges,
	type OnDestroy,
	type OnInit,
} from '@angular/core';
import {
	buildingBlock,
	createBuildingBlockSession,
	type BuildingBlockOptions,
} from '../../bindings/src/building-blocks';
@Injectable()
export class ViewerBuildingBlocks {
	private readonly session = createBuildingBlockSession();
	readonly controller = this.session.controller;
	private readonly current = signal(this.controller.state);
	readonly state = this.current.asReadonly();
	private readonly locale = signal<string | undefined>(undefined);
	readonly toolbarProps = computed(() => {
		this.state();
		return this.session.toolbar(this.locale());
	});
	readonly loading = computed(() => this.state().loading);
	readonly error = computed(() => this.state().error ?? this.state().edit.error);
	readonly mode = computed(() => (this.state().canEdit ? ('edit' as const) : ('view' as const)));
	readonly documentChange = new EventEmitter<NonNullable<typeof this.controller.state.document>>();
	readonly dirtyChange = new EventEmitter<boolean>();
	readonly documentError = new EventEmitter<Error>();
	readonly canvasProps = { mount: (node: HTMLElement | null) => this.session.mount(node) };
	readonly inspectorProps = { controller: this.controller };
	constructor() {
		const stop = this.controller.subscribe((state) => this.current.set(state));
		inject(DestroyRef).onDestroy(() => {
			stop();
			this.session.destroy();
		});
	}
	update(options: BuildingBlockOptions) {
		this.locale.set(options.locale);
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
@Component({
	selector: 'visio-building-blocks',
	standalone: true,
	template: '<ng-content />',
	providers: [ViewerBuildingBlocks],
})
export class ViewerBuildingBlocksComponent implements OnChanges, OnInit {
	readonly blocks = inject(ViewerBuildingBlocks);
	@Input() content?: BuildingBlockOptions['content'];
	@Input() canEdit = true;
	@Input() locale?: string;
	@Output() readonly documentChange = this.blocks.documentChange;
	@Output() readonly dirtyChange = this.blocks.dirtyChange;
	@Output() readonly documentError = this.blocks.documentError;
	ngOnInit() {
		this.ngOnChanges();
	}
	ngOnChanges() {
		this.blocks.update({ content: this.content, canEdit: this.canEdit, locale: this.locale });
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
		<button
			type="button"
			[disabled]="!blocks.toolbarProps().canUndo"
			(click)="blocks.toolbarProps().onUndo()"
		>
			{{ blocks.toolbarProps().labels['Undo'] }}</button
		><button
			type="button"
			[disabled]="!blocks.toolbarProps().canRedo"
			(click)="blocks.toolbarProps().onRedo()"
		>
			{{ blocks.toolbarProps().labels['Redo'] }}</button
		><button type="button" (click)="blocks.toolbarProps().onFit()">
			{{ blocks.toolbarProps().labels['Fit to Window'] }}
		</button>
	</div>`,
})
export class ToolbarComponent {
	readonly blocks = inject(ViewerBuildingBlocks);
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
export class PageNavigationComponent {
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
export type { BuildingBlockOptions, DiagramToolbarProps } from '../../bindings/src/building-blocks';
