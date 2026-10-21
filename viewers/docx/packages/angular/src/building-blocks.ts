import {
	Component,
	DestroyRef,
	ElementRef,
	EventEmitter,
	Injectable,
	Input,
	effect,
	inject,
	signal,
	computed,
	untracked,
	type AfterViewInit,
	type OnChanges,
	type OnDestroy,
} from '@angular/core';
import {
	createBuildingBlockSession,
	prepareWordToolbar,
	type BuildingBlockOptions,
	type CanvasProps,
	type WordToolbarProps,
} from '../../bindings/src/building-blocks';
import type { PageChangeDetail } from 'docx-web-component';
import type { DocumentModel } from 'docx-core';
export type {
	BuildingBlockOptions,
	CanvasProps,
	WordToolbarProps,
} from '../../bindings/src/building-blocks';
/** Provide per editor, so Angular owns the composition's state and lifetime. */
@Injectable()
export class ViewerBuildingBlocks {
	private readonly session = createBuildingBlockSession();
	private readonly state = signal(this.session.state);
	readonly toolbarProps = computed(() => this.state().toolbarProps);
	readonly loading = computed(() => this.state().loading);
	readonly error = computed(() => this.state().error);
	readonly mode = computed(() => this.state().mode);
	readonly canvasProps = this.session.canvasProps;
	readonly handle = this.session.handle;
	readonly documentChange = new EventEmitter<DocumentModel>();
	readonly documentError = new EventEmitter<Error>();
	readonly dirtyChange = new EventEmitter<boolean>();
	readonly pageChange = new EventEmitter<PageChangeDetail>();
	constructor() {
		const stop = this.session.subscribe((state) => this.state.set(state));
		inject(DestroyRef).onDestroy(() => {
			stop();
			this.session.destroy();
			this.documentChange.complete();
			this.documentError.complete();
			this.dirtyChange.complete();
			this.pageChange.complete();
		});
		this.update({});
	}
	update(options: BuildingBlockOptions) {
		this.session.update({
			...options,
			onDocumentChange: (model) => {
				this.documentChange.emit(model);
				options.onDocumentChange?.(model);
			},
			onDocumentError: (error) => {
				this.documentError.emit(error);
				options.onDocumentError?.(error);
			},
			onPageChange: (page) => {
				this.pageChange.emit(page);
				options.onPageChange?.(page);
			},
			onDirtyChange: (dirty) => {
				this.dirtyChange.emit(dirty);
				options.onDirtyChange?.(dirty);
			},
		});
	}
}
export function createViewerBuildingBlocks(options: () => BuildingBlockOptions = () => ({})) {
	const blocks = inject(ViewerBuildingBlocks);
	effect(() => {
		const next = options();
		untracked(() => blocks.update(next));
	});
	return blocks;
}
@Component({ selector: 'word-document-canvas', standalone: true, template: '' })
export class DocumentCanvasComponent implements AfterViewInit, OnChanges, OnDestroy {
	@Input({ required: true }) canvasProps!: CanvasProps;
	private readonly host = inject<ElementRef<HTMLElement>>(ElementRef);
	private mounted: CanvasProps | undefined;
	ngAfterViewInit() {
		this.attach();
	}
	ngOnChanges() {
		if (this.mounted) this.attach();
	}
	private attach() {
		if (this.mounted?.mount === this.canvasProps.mount) {
			this.mounted = this.canvasProps;
			return;
		}
		this.mounted?.mount(null);
		this.mounted = this.canvasProps;
		this.mounted.mount(this.host.nativeElement);
	}
	ngOnDestroy() {
		this.mounted?.mount(null);
		this.mounted = undefined;
	}
}
@Component({
	selector: 'word-toolbar',
	standalone: true,
	template: ` <div role="toolbar">
		<button
			type="button"
			[disabled]="!toolbarProps.canEdit"
			[attr.aria-pressed]="toolbarProps.bold === null ? 'mixed' : toolbarProps.bold"
			(mousedown)="$event.preventDefault()"
			(click)="toolbarProps.onBold()"
		>
			{{ toolbarProps.labels['Bold'] }}
		</button>
		<button
			type="button"
			[disabled]="!toolbarProps.canEdit"
			[attr.aria-pressed]="toolbarProps.italic === null ? 'mixed' : toolbarProps.italic"
			(mousedown)="$event.preventDefault()"
			(click)="toolbarProps.onItalic()"
		>
			{{ toolbarProps.labels['Italic'] }}
		</button>
		<button
			type="button"
			[disabled]="!toolbarProps.canEdit"
			[attr.aria-pressed]="toolbarProps.underline === null ? 'mixed' : toolbarProps.underline"
			(mousedown)="$event.preventDefault()"
			(click)="toolbarProps.onUnderline()"
		>
			{{ toolbarProps.labels['Underline'] }}
		</button>
		<button type="button" [disabled]="!toolbarProps.canUndo" (click)="toolbarProps.onUndo()">
			{{ toolbarProps.labels['Undo'] }}
		</button>
		<button type="button" [disabled]="!toolbarProps.canRedo" (click)="toolbarProps.onRedo()">
			{{ toolbarProps.labels['Redo'] }}
		</button>
	</div>`,
})
export class ToolbarComponent implements AfterViewInit {
	private readonly host = inject<ElementRef<HTMLElement>>(ElementRef);
	ngAfterViewInit() {
		prepareWordToolbar(this.host.nativeElement.querySelector<HTMLElement>('[role="toolbar"]'));
	}
	@Input({ required: true }) toolbarProps!: WordToolbarProps;
}
