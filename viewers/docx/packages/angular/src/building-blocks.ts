import {
	Component,
	Directive,
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
	type WordNavigationProps,
	type WordStatusProps,
	type WordToolbarProps,
} from '../../bindings/src/building-blocks';
import {
	navigationControls,
	statusControls,
	toolbarControls,
	type ControlGroup,
} from '../../bindings/src/controls';
import type { PageChangeDetail } from 'docx-web-component';
import type { DocumentModel } from 'docx-core';
export type {
	BuildingBlockOptions,
	CanvasProps,
	WordNavigationProps,
	WordStatusProps,
	WordToolbarProps,
} from '../../bindings/src/building-blocks';
/** Provide per editor, so Angular owns the composition's state and lifetime. */
@Injectable()
export class ViewerBuildingBlocks {
	private readonly session = createBuildingBlockSession();
	private readonly state = signal(this.session.state);
	readonly toolbarProps = computed(() => this.state().toolbarProps);
	readonly navigationProps = computed(() => this.state().navigationProps);
	readonly statusProps = computed(() => this.state().statusProps);
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
const CONTROLS_TEMPLATE = `<div [attr.role]="role">
	@for (group of groups; track $index) {
		<span role="group">
			@for (control of group; track control.id) {
				@switch (control.kind) {
					@case ('text') {
						<span role="status">{{ $any(control).text }}</span>
					}
					@case ('select') {
						<select
							[attr.aria-label]="control.label"
							[disabled]="control.disabled"
							[value]="$any(control).value"
							(change)="$any(control).change($any($event.target).value)"
						>
							@for (option of $any(control).options; track option.value) {
								<option [value]="option.value" [selected]="option.value === $any(control).value">
									{{ option.label }}
								</option>
							}
						</select>
					}
					@default {
						<button
							type="button"
							[attr.aria-label]="control.label"
							[disabled]="control.disabled"
							[attr.aria-pressed]="control.kind === 'toggle' ? $any(control).pressed : null"
							(mousedown)="$event.preventDefault()"
							(click)="$any(control).run()"
						>
							{{ control.label }}
						</button>
					}
				}
			}
		</span>
	}
</div>`;
/** Draws control groups from props; nodes are tracked by control id so focus survives updates. */
@Directive()
abstract class ControlsComponent implements AfterViewInit {
	private readonly host = inject<ElementRef<HTMLElement>>(ElementRef);
	protected abstract readonly role: string;
	protected abstract get groups(): readonly ControlGroup[];
	ngAfterViewInit() {
		prepareWordToolbar(this.host.nativeElement.querySelector<HTMLElement>('div'));
	}
}
/** Optional controls: history, font, formatting, alignment, lists and enabled host workflows. */
@Component({ selector: 'word-toolbar', standalone: true, template: CONTROLS_TEMPLATE })
export class ToolbarComponent extends ControlsComponent {
	@Input({ required: true }) toolbarProps!: WordToolbarProps;
	protected readonly role = 'toolbar';
	protected get groups() {
		return toolbarControls(this.toolbarProps);
	}
}
/** Optional page navigation for `navigationProps`. Pages exist in Print Layout only. */
@Component({ selector: 'word-page-navigation', standalone: true, template: CONTROLS_TEMPLATE })
export class PageNavigationComponent extends ControlsComponent {
	@Input({ required: true }) navigationProps!: WordNavigationProps;
	protected readonly role = 'group';
	protected get groups() {
		return navigationControls(this.navigationProps);
	}
}
/** Optional word count and zoom for `statusProps`. */
@Component({ selector: 'word-status-bar', standalone: true, template: CONTROLS_TEMPLATE })
export class StatusBarComponent extends ControlsComponent {
	@Input({ required: true }) statusProps!: WordStatusProps;
	protected readonly role = 'group';
	protected get groups() {
		return statusControls(this.statusProps);
	}
}
