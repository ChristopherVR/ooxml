/**
 * inspector-fields.component.ts: the Angular binding's one generic field
 * renderer. It draws any list of shared inspector section descriptors (a
 * toggle per section, then its fields); values, bounds, labels and patches all
 * come from the descriptors (`resolveInspectorSections`, `inspectorFieldPatch`,
 * `inspectorTogglePatch` in `ooxml-ui/pptx`), so this component only maps them
 * onto its template. A colour commit also pushes into the "Recent colours"
 * list.
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
import { inspectorFieldPatch, inspectorTogglePatch, resolveInspectorSections } from 'ooxml-ui/pptx';
import type {
	InspectorFieldValue,
	InspectorSectionDescriptor,
	InspectorSectionView,
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
												(change)="onColor(section.id, field.id, $event)"
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

		.pptx-ng-fields__fields {
			display: flex;
			flex-wrap: wrap;
			align-items: center;
			gap: 0.3rem;
			padding: 0.35rem 0 0 1.25rem;
		}

		.pptx-ng-fields__field {
			display: inline-flex;
			align-items: center;
			gap: 0.3rem;
		}

		.pptx-ng-fields__wide {
			flex-basis: 100%;
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

	protected onColor(sectionId: string, fieldId: string, event: Event): void {
		const raw = rawValueOf(event);
		if (typeof raw !== 'string') {
			return;
		}
		const patch = inspectorFieldPatch(this.sections(), sectionId, fieldId, this.target(), raw);
		this.emit(patch);
		if (patch !== undefined) {
			this.recentColors?.push(raw);
		}
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
