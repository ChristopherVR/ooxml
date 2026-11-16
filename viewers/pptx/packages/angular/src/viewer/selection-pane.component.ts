/**
 * selection-pane.component.ts: Side panel listing all elements on the active slide.
 *
 * Selector: `pptx-selection-pane`
 *
 * Displays elements in reverse z-order (topmost first = last in the elements array).
 * Each row shows a type icon, the shared row label (name, else text, else the
 * translated type and position), visibility toggle, and z-order nudge buttons.
 * Clicking a row selects that element; clicking the eye icon toggles its `hidden`
 * flag; the up/down arrows bring it forward or send it backward; dragging a row
 * reorders it; double-clicking the name label edits it inline (Enter/blur commit,
 * Escape cancels). The rename box and the drag are the shared selection pane
 * controller; every deck change is surfaced as an output.
 */

import {
	ChangeDetectionStrategy,
	Component,
	DestroyRef,
	computed,
	inject,
	input,
	output,
} from '@angular/core';
import { LucideArrowDown, LucideArrowUp, LucideEye, LucideEyeOff } from '@lucide/angular';
import { TranslatePipe, TranslateService } from '@ngx-translate/core';
import {
	createSelectionPaneController,
	formatSelectionPaneLabel,
	selectionPaneRows,
} from 'ooxml-ui/pptx';
import type { PptxElement } from 'pptx-viewer-core';

import { controllerSignal } from './controller-signal';
import { elementIcon } from './selection-pane-helpers';
import { translationsSignal } from './translations-signal';

/** Payload of `renameElement`: an empty `name` clears the element's name. */
export interface SelectionPaneRename {
	id: string;
	name: string;
}

@Component({
	selector: 'pptx-selection-pane',
	standalone: true,
	changeDetection: ChangeDetectionStrategy.OnPush,
	imports: [TranslatePipe, LucideEye, LucideEyeOff, LucideArrowUp, LucideArrowDown],
	template: `
		<aside
			class="pptx-ng-sel-pane"
			data-pptx-selection-pane
			[attr.aria-label]="'pptx.selectionPane.title' | translate"
		>
			<header class="pptx-ng-sel-pane__header">
				<h2 class="pptx-ng-sel-pane__title">{{ 'pptx.selectionPane.title' | translate }}</h2>
				<span class="pptx-ng-sel-pane__count">{{ elements().length }}</span>
			</header>

			@if (rows().length > 0) {
				<ul class="pptx-ng-sel-pane__list" role="listbox" aria-multiselectable="true">
					@for (row of rows(); track row.element.id) {
						@let el = row.element;
						<li
							class="pptx-ng-sel-pane__row"
							[class.pptx-ng-sel-pane__row--selected]="isSelected(el.id)"
							[class.pptx-ng-sel-pane__row--hidden]="el.hidden"
							role="option"
							[attr.aria-selected]="isSelected(el.id)"
							[attr.draggable]="canEdit() && editingId() !== el.id ? 'true' : 'false'"
							[class.pptx-ng-sel-pane__row--drop]="dragOverIndex() === row.index"
							(click)="selectElement.emit(el.id)"
							(dragstart)="pane.dragStart(row.index)"
							(dragover)="onDragOver($event, row.index)"
							(drop)="pane.drop(row.index)"
							(dragend)="pane.dragEnd()"
						>
							<span class="pptx-ng-sel-pane__icon" [attr.title]="el.type" aria-hidden="true">{{
								typeIcon(el.type)
							}}</span>

							<span
								class="pptx-ng-sel-pane__label"
								data-pptx-selection-name
								[title]="row.name"
								(dblclick)="startRename(el.id, row.name)"
							>
								@if (editingId() === el.id) {
									<input
										type="text"
										class="pptx-ng-sel-pane__rename"
										autofocus
										[value]="paneState().renameDraft"
										[attr.aria-label]="'pptx.selectionPane.renameElement' | translate"
										(click)="$event.stopPropagation()"
										(input)="setRenameDraft($event)"
										(keydown.enter)="pane.commitRename()"
										(keydown.escape)="cancelRename($event)"
										(blur)="pane.commitRename()"
									/>
								} @else {
									<ng-container>{{ row.name }}</ng-container>
								}
							</span>

							<div class="pptx-ng-sel-pane__actions" (click)="$event.stopPropagation()">
								<button
									type="button"
									class="pptx-ng-sel-pane__btn"
									[attr.aria-label]="
										(el.hidden
											? 'pptx.selectionPane.showElement'
											: 'pptx.selectionPane.hideElement'
										) | translate
									"
									[title]="
										(el.hidden ? 'pptx.selectionPane.show' : 'pptx.selectionPane.hide') | translate
									"
									(click)="toggleHidden.emit(el.id)"
								>
									@if (el.hidden) {
										<svg lucideEyeOff class="h-3.5 w-3.5"></svg>
									} @else {
										<svg lucideEye class="h-3.5 w-3.5"></svg>
									}
								</button>
								<button
									type="button"
									class="pptx-ng-sel-pane__btn"
									[attr.aria-label]="'pptx.arrange.bringForward' | translate"
									[title]="'pptx.arrange.bringForward' | translate"
									(click)="bringForward.emit(el.id)"
								>
									<svg lucideArrowUp class="h-3.5 w-3.5"></svg>
								</button>
								<button
									type="button"
									class="pptx-ng-sel-pane__btn"
									[attr.aria-label]="'pptx.arrange.sendBackward' | translate"
									[title]="'pptx.arrange.sendBackward' | translate"
									(click)="sendBackward.emit(el.id)"
								>
									<svg lucideArrowDown class="h-3.5 w-3.5"></svg>
								</button>
							</div>
						</li>
					}
				</ul>
			} @else {
				<p class="pptx-ng-sel-pane__empty">{{ 'pptx.selectionPane.empty' | translate }}</p>
			}
		</aside>
	`,
	styleUrl: './selection-pane.component.css',
})
export class SelectionPaneComponent {
	private readonly translate = inject(TranslateService);
	private readonly translations = translationsSignal(this.translate);

	/** All elements on the active slide. */
	readonly elements = input<readonly PptxElement[]>([]);
	/** Currently selected element ids. */
	readonly selectedIds = input<readonly string[]>([]);
	/** Whether the host allows edits: renaming and dragging are refused otherwise. */
	readonly canEdit = input<boolean>(true);

	/** Emits the id of the element the user clicked. */
	readonly selectElement = output<string>();
	/** Emits the id of the element to bring one step forward in z-order. */
	readonly bringForward = output<string>();
	/** Emits the id of the element to send one step backward in z-order. */
	readonly sendBackward = output<string>();
	/** Emits the id of the element whose hidden flag should be toggled. */
	readonly toggleHidden = output<string>();
	/** Emits the rename commit; an empty `name` clears the element's name. */
	readonly renameElement = output<SelectionPaneRename>();
	/** Emits a drag reorder, as stacking positions (0 is the back-most). */
	readonly reorderElement = output<{ from: number; to: number }>();

	/** The shared rename box and drag state. */
	protected readonly pane = createSelectionPaneController({
		canEdit: () => this.canEdit(),
		onRename: (id, name) => this.renameElement.emit({ id, name }),
		onReorder: (from, to) => this.reorderElement.emit({ from, to }),
	});
	protected readonly paneState = controllerSignal(this.pane, inject(DestroyRef));
	protected readonly editingId = computed(() => this.paneState().renamingId);
	protected readonly dragOverIndex = computed(() => this.paneState().dragOverIndex);

	/** The rows, top-most first, with the shared label. */
	protected readonly rows = computed(() => {
		this.translations();
		return selectionPaneRows(this.elements()).map((row) => ({
			...row,
			name: formatSelectionPaneLabel(row.label, (key, params) =>
				this.translate.instant(key, params),
			),
		}));
	});

	/** Double-click on the name label: open the inline rename input. */
	protected startRename(id: string, name: string): void {
		this.pane.startRename(id, name);
	}

	protected setRenameDraft(event: Event): void {
		this.pane.setRenameDraft((event.target as HTMLInputElement).value);
	}

	/** Escape: drop the edit without committing. */
	protected cancelRename(event: Event): void {
		event.stopPropagation();
		this.pane.cancelRename();
	}

	protected onDragOver(event: DragEvent, index: number): void {
		event.preventDefault();
		this.pane.dragOver(index);
	}

	protected isSelected(id: string): boolean {
		return this.selectedIds().includes(id);
	}

	protected typeIcon(type: string): string {
		return elementIcon(type);
	}
}
