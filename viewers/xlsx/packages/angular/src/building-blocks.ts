import { prepareBlock } from '../../bindings/src/block-theme';
import {
	Component,
	EventEmitter,
	DestroyRef,
	ElementRef,
	Injectable,
	Input,
	computed,
	effect,
	inject,
	signal,
	untracked,
	type AfterViewInit,
	type OnDestroy,
	type OnChanges,
} from '@angular/core';
import {
	createBuildingBlockSession,
	emptyBuildingBlockState,
	toolbarButtons,
	selectReference,
	type BuildingBlockOptions,
	type SpreadsheetToolbarProps,
} from '../../bindings/src/building-blocks';
export type {
	BuildingBlockOptions,
	SpreadsheetToolbarProps,
} from '../../bindings/src/building-blocks';
/** Provide per editor host, then inject into custom controls. Effects and subscriptions follow its injector. */
@Injectable()
export class SpreadsheetBuildingBlocks {
	readonly workbookChange = new EventEmitter<
		Parameters<NonNullable<BuildingBlockOptions['onWorkbookChange']>>[0]
	>();
	readonly workbookError = new EventEmitter<Error>();
	readonly selectionChange = new EventEmitter<
		Parameters<NonNullable<BuildingBlockOptions['onSelectionChange']>>[0]
	>();
	readonly dirtyChange = new EventEmitter<boolean>();
	readonly options = signal<BuildingBlockOptions>({});
	private readonly state = signal(emptyBuildingBlockState());
	private readonly session = createBuildingBlockSession((next) => this.state.set(next));
	readonly toolbarProps = computed(() => this.state().toolbarProps);
	readonly loading = computed(() => this.state().loading);
	readonly error = computed(() => this.state().error);
	readonly mode = computed(() => this.state().mode);
	readonly handle = this.session.handle;
	readonly canvasProps = { mount: (host: HTMLElement | null) => this.session.mount(host) };
	constructor() {
		const effective = (options: BuildingBlockOptions): BuildingBlockOptions => ({
			...options,
			onWorkbookChange: (value) => {
				this.workbookChange.emit(value);
				this.options().onWorkbookChange?.(value);
			},
			onWorkbookError: (value) => {
				this.workbookError.emit(value);
				this.options().onWorkbookError?.(value);
			},
			onSelectionChange: (value) => {
				this.selectionChange.emit(value);
				this.options().onSelectionChange?.(value);
			},
			onDirtyChange: (value) => {
				this.dirtyChange.emit(value);
				this.options().onDirtyChange?.(value);
			},
		});
		this.session.update(effective(this.options()));
		effect(() => {
			const options = this.options();
			untracked(() => this.session.update(effective(options)));
		});
		inject(DestroyRef).onDestroy(() => this.session.destroy());
	}
}
@Component({ selector: 'spreadsheet-canvas', standalone: true, template: '' })
export class SpreadsheetCanvasComponent implements AfterViewInit, OnDestroy, OnChanges {
	@Input({ required: true }) mount!: (host: HTMLElement | null) => void;
	private readonly host = inject<ElementRef<HTMLElement>>(ElementRef);
	private mounted: ((host: HTMLElement | null) => void) | undefined;
	ngAfterViewInit() {
		this.mounted = this.mount;
		this.mounted(this.host.nativeElement);
	}
	ngOnChanges() {
		if (this.mounted && this.mounted !== this.mount) {
			this.mounted(null);
			this.mounted = this.mount;
			this.mounted(this.host.nativeElement);
		}
	}
	ngOnDestroy() {
		this.mounted?.(null);
		this.mounted = undefined;
	}
}
@Component({
	selector: 'spreadsheet-toolbar',
	standalone: true,
	template: `<div role="toolbar">
		@for (button of buttons(); track button.label) {
			<button
				type="button"
				[disabled]="button.disabled"
				[attr.aria-pressed]="button.pressed === null ? 'mixed' : button.pressed"
				(mousedown)="$event.preventDefault()"
				(click)="button.run()"
			>
				{{ button.label }}
			</button>
		}
	</div>`,
})
export class ToolbarComponent implements AfterViewInit {
	private readonly themeHost = inject<ElementRef<HTMLElement>>(ElementRef);
	ngAfterViewInit() {
		prepareBlock(this.themeHost.nativeElement, 'toolbar');
	}
	@Input({ required: true }) toolbarProps!: SpreadsheetToolbarProps;
	buttons() {
		return toolbarButtons(this.toolbarProps);
	}
}
@Component({
	selector: 'spreadsheet-formula-bar',
	standalone: true,
	template: `<div>
		<input
			[attr.aria-label]="toolbarProps.labels['Name Box']"
			[value]="toolbarProps.selection"
			[attr.aria-invalid]="invalidName"
			(input)="invalidName = false"
			(keydown.enter)="select($event)"
		/>
		<input
			[attr.aria-label]="toolbarProps.labels['Formula Bar']"
			[value]="toolbarProps.formula"
			[readOnly]="!toolbarProps.canEdit"
			[attr.aria-invalid]="invalid"
			(input)="draft($event)"
			(keydown.enter)="commit($event)"
			(keydown.escape)="cancel($event)"
		/>
	</div>`,
})
export class FormulaBarComponent implements AfterViewInit {
	private readonly themeHost = inject<ElementRef<HTMLElement>>(ElementRef);
	ngAfterViewInit() {
		prepareBlock(this.themeHost.nativeElement, 'formula-bar');
	}
	@Input({ required: true }) toolbarProps!: SpreadsheetToolbarProps;
	invalid = false;
	invalidName = false;
	select(e: Event) {
		e.preventDefault();
		this.invalidName = !selectReference(this.toolbarProps, (e.target as HTMLInputElement).value);
	}
	draft(e: Event) {
		const input = e.target as HTMLInputElement;
		this.invalid = false;
		this.toolbarProps.onFormulaInput(input.value, input.selectionEnd ?? undefined);
	}
	commit(e: Event) {
		e.preventDefault();
		this.invalid = !this.toolbarProps.onFormulaChange((e.target as HTMLInputElement).value);
	}
	cancel(e: Event) {
		e.preventDefault();
		this.toolbarProps.onCancelEdit();
		this.invalid = false;
	}
}
@Component({
	selector: 'spreadsheet-sheet-tabs',
	standalone: true,
	template: `<div role="tablist" [attr.aria-label]="toolbarProps.labels['Sheet']">
		@for (sheet of toolbarProps.sheets; track sheet.index) {
			@if (sheet.state === 'visible') {
				<button
					type="button"
					role="tab"
					[attr.aria-selected]="sheet.index === toolbarProps.activeSheet"
					(click)="toolbarProps.onSheetChange(sheet.index)"
				>
					{{ sheet.name }}
				</button>
			}
		}
	</div>`,
})
export class SheetTabsComponent implements AfterViewInit {
	private readonly themeHost = inject<ElementRef<HTMLElement>>(ElementRef);
	ngAfterViewInit() {
		prepareBlock(this.themeHost.nativeElement, 'tabs');
	}
	@Input({ required: true }) toolbarProps!: SpreadsheetToolbarProps;
}

export { SpreadsheetBuildingBlocks as ViewerBuildingBlocks };
