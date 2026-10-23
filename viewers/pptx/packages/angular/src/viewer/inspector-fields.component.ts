/**
 * inspector-fields.component.ts: the Angular binding's one generic field
 * renderer. It draws any list of shared inspector section descriptors (a
 * toggle per section, then its fields); values, bounds, labels and patches all
 * come from the descriptors (`resolveInspectorSections`, `inspectorFieldPatch`,
 * `inspectorTogglePatch` in `ooxml-ui/pptx`), so this component only maps them
 * onto its template. A colour field previews while the user drags through
 * the native picker (`preview`, no undo step) and commits once when the pick
 * settles (`patch`, plus the "Recent colours" list), through the shared
 * `createLiveColorPatchInput`.
 *
 * Selector: `pptx-inspector-fields`
 *
 * @module viewer/inspector-fields
 */
import {
	ChangeDetectionStrategy,
	Component,
	computed,
	CUSTOM_ELEMENTS_SCHEMA,
	inject,
	input,
	output,
} from '@angular/core';
import { TranslatePipe } from '@ngx-translate/core';
import {
	createLiveColorPatchInput,
	handleLiveColorEvent,
	inspectorFieldPatch,
	inspectorTogglePatch,
	resolveInspectorSections,
} from 'ooxml-ui/pptx';
import type {
	InspectorFieldValue,
	InspectorSectionDescriptor,
	InspectorSectionView,
	LiveColorInput,
} from 'ooxml-ui/pptx';

import { isCheckboxControl, isSelectControl } from './control-event-targets';
import { RecentColorsService } from './recent-colors.service';

@Component({
	schemas: [CUSTOM_ELEMENTS_SCHEMA],
	selector: 'pptx-inspector-fields',
	standalone: true,
	imports: [TranslatePipe],
	changeDetection: ChangeDetectionStrategy.OnPush,
	template: `
		<div class="pptx-ng-fields">
			@for (section of views(); track section.id) {
				<section class="pptx-ng-fields__section" [attr.data-section]="section.id">
					@if (section.toggleable) {
						<label class="pptx-ng-fields__toggle-row">
							<pptx-ui-checkbox
								class="pptx-ng-fields__checkbox"
								[attr.data-testid]="section.toggleTestId"
								[checked]="section.on"
								[disabled]="disabled()"
								(change)="onToggle(section.id, $event)"
							></pptx-ui-checkbox>
							<span class="pptx-ng-fields__section-title">{{
								section.labelKey ?? '' | translate
							}}</span>
						</label>
					} @else if (section.labelKey) {
						<span class="pptx-ng-fields__section-title">{{ section.labelKey | translate }}</span>
					}
					@if (section.fields.length) {
						<div class="pptx-ng-fields__fields">
							@for (field of section.fields; track field.id) {
								@if (field.control === 'checkbox') {
									<label class="pptx-ng-fields__toggle-row pptx-ng-fields__wide">
										<pptx-ui-checkbox
											class="pptx-ng-fields__checkbox"
											[attr.data-testid]="field.testId"
											[checked]="field.value === true"
											[disabled]="disabled()"
											(change)="onField(section.id, field.id, $event)"
										></pptx-ui-checkbox>
										<span>{{ field.labelKey | translate }}</span>
									</label>
								} @else {
									<label class="pptx-ng-fields__field" [class.pptx-ng-fields__wide]="field.wide">
										<span class="pptx-ng-fields__label">{{ field.labelKey | translate }}</span>
										@if (field.control === 'select') {
											<pptx-ui-select
												class="pptx-ng-fields__input"
												[attr.data-testid]="field.testId"
												[attr.aria-label]="field.labelKey | translate"
												[value]="'' + field.value"
												[disabled]="disabled()"
												(change)="onField(section.id, field.id, $event)"
											>
												@for (option of field.options ?? []; track option.value) {
													<option
														[value]="option.value"
														[selected]="option.value === '' + field.value"
													>
														{{ option.labelKey | translate }}
													</option>
												}
											</pptx-ui-select>
										} @else if (field.control === 'color') {
											<input
												class="pptx-ng-fields__color"
												type="color"
												[attr.data-testid]="field.testId"
												[attr.aria-label]="field.labelKey | translate"
												[value]="field.value"
												[disabled]="disabled()"
												(input)="onColor(section.id, field.id, $event)"
												(change)="onColor(section.id, field.id, $event)"
												(blur)="onColor(section.id, field.id, $event)"
											/>
										} @else {
											<input
												class="pptx-ng-fields__input pptx-ng-fields__input--number"
												[type]="field.control"
												inputmode="decimal"
												[attr.data-testid]="field.testId"
												[attr.aria-label]="field.labelKey | translate"
												[attr.min]="field.min"
												[attr.max]="field.max"
												[attr.step]="field.step"
												[value]="field.value"
												[disabled]="disabled()"
												(change)="onField(section.id, field.id, $event)"
											/>
										}
									</label>
								}
							}
						</div>
					}
				</section>
			}
		</div>
	`,
	styles: `
		.pptx-ng-fields {
			display: flex;
			flex-direction: column;
			padding: 0.5rem;
			font-size: 12px;
			color: var(--pptx-inspector-fg, #e0e0e0);
		}

		.pptx-ng-fields__section {
			padding: 0.35rem 0;
			border-bottom: 1px solid var(--pptx-inspector-border, #333);
		}

		.pptx-ng-fields__section:last-child {
			border-bottom: none;
		}

		.pptx-ng-fields__section-title {
			font-size: 11px;
			font-weight: 500;
		}

		.pptx-ng-fields__toggle-row {
			display: flex;
			align-items: center;
			gap: 0.4rem;
			cursor: pointer;
		}

		.pptx-ng-fields__checkbox {
			cursor: pointer;
		}

		/*
		 * A two-column grid, as every binding's generic renderer draws it: paired
		 * fields (X/Y, W/H) share a row whatever the pane's width. A wrapping flex
		 * row let a narrow phone sheet push H under W.
		 */
		.pptx-ng-fields__fields {
			display: grid;
			grid-template-columns: minmax(0, 1fr) minmax(0, 1fr);
			align-items: center;
			gap: 0.3rem;
			padding: 0.35rem 0 0 1.25rem;
		}

		.pptx-ng-fields__field {
			display: flex;
			align-items: center;
			gap: 0.3rem;
			min-width: 0;
		}

		.pptx-ng-fields__wide {
			grid-column: 1 / -1;
		}

		.pptx-ng-fields__label {
			font-size: 10px;
			color: var(--pptx-inspector-muted, #888);
			min-width: 36px;
			text-align: right;
			flex-shrink: 0;
		}

		.pptx-ng-fields__input {
			background: var(--pptx-inspector-input-bg, #2d2d2d);
			border: 1px solid var(--pptx-inspector-border, #444);
			color: inherit;
			border-radius: 3px;
			padding: 2px 4px;
			font-size: 12px;
		}

		.pptx-ng-fields__input--number {
			width: 56px;
			min-width: 0;
			flex: 0 1 auto;
			text-align: right;
		}

		.pptx-ng-fields__color {
			width: 32px;
			height: 22px;
			border: 1px solid var(--pptx-inspector-border, #444);
			border-radius: 3px;
			padding: 1px;
			cursor: pointer;
			background: transparent;
			flex-shrink: 0;
		}

		/* Touch: 44px targets, 16px text (no iOS focus zoom), inputs fill their column. */
		@media (pointer: coarse), (max-width: 767px) {
			.pptx-ng-fields__input {
				min-height: 44px;
				font-size: 16px;
			}

			.pptx-ng-fields__input--number {
				width: 100%;
				flex: 1 1 auto;
			}
		}
	`,
})
export class InspectorFieldsComponent<T, P> {
	/** The section descriptors to draw (from `ooxml-ui/pptx`). */
	readonly sections = input.required<readonly InspectorSectionDescriptor<T, P>[]>();
	/** The value the descriptors read. */
	readonly target = input.required<T>();
	readonly disabled = input(false);

	/** Emits the patch a control change produces. */
	readonly patch = output<P>();
	/**
	 * Emits a colour field's live patch while its picker is dragged: apply it
	 * WITHOUT an undo step. The drag's one `patch` follows when it settles.
	 */
	readonly preview = output<P>();

	/** Optional: absent in a standalone unit test with no viewer-level DI tree. */
	private readonly recentColors = inject(RecentColorsService, { optional: true });

	protected readonly views = computed<InspectorSectionView[]>(() =>
		resolveInspectorSections(this.sections(), this.target()),
	);

	protected onToggle(sectionId: string, event: Event): void {
		if (!isCheckboxControl(event.target)) {
			return;
		}
		this.emit(
			inspectorTogglePatch(this.sections(), sectionId, this.target(), event.target.checked),
		);
	}

	protected onField(sectionId: string, fieldId: string, event: Event): void {
		const raw = rawValueOf(event);
		if (raw === null) {
			return;
		}
		this.emit(inspectorFieldPatch(this.sections(), sectionId, fieldId, this.target(), raw));
	}

	/** One live colour controller per colour field, created on its first event. */
	private readonly colorInputs = new Map<string, LiveColorInput>();

	protected onColor(sectionId: string, fieldId: string, event: Event): void {
		const key = `${sectionId}.${fieldId}`;
		let control = this.colorInputs.get(key);
		if (!control) {
			control = createLiveColorPatchInput<T, P & object>({
				read: () => {
					const field = this.views()
						.find((section) => section.id === sectionId)
						?.fields.find((candidate) => candidate.id === fieldId);
					return field && { key, color: String(field.value), state: this.target() };
				},
				patch: (target, hex) =>
					inspectorFieldPatch(this.sections(), sectionId, fieldId, target, hex) as
						| (P & object)
						| undefined,
				preview: (patch) => this.preview.emit(patch),
				commit: (patch) => this.patch.emit(patch),
				remember: (hex) => this.recentColors?.push(hex),
			});
			this.colorInputs.set(key, control);
		}
		handleLiveColorEvent(control, event);
	}

	private emit(patch: P | undefined): void {
		if (patch !== undefined) {
			this.patch.emit(patch);
		}
	}
}

/** The value a field control reports: `checked` for a checkbox, else its `value`. */
function rawValueOf(event: Event): InspectorFieldValue | null {
	const target = event.target;
	if (target instanceof HTMLElement && target.localName === 'pptx-ui-checkbox') {
		return isCheckboxControl(target) ? target.checked : null;
	}
	if (target instanceof HTMLInputElement) {
		return target.type === 'checkbox' ? target.checked : target.value;
	}
	return isSelectControl(target) ? target.value : null;
}
