/**
 * inspector-fields.component.ts: the Angular binding's one generic field
 * renderer. It draws any list of shared inspector section descriptors (a
 * toggle per section, then its fields); values, bounds, labels and patches all
 * come from the descriptors (`resolveInspectorSections`, `inspectorFieldPatch`,
 * `inspectorTogglePatch` in `ooxml-ui/pptx`), so this component only maps them
 * onto its template. A colour field previews while the user drags through
 * the native picker (`preview`, no undo step) and commits once when the pick
 * settles (`patch`, plus the "Recent colours" list), through the shared
 * `createInspectorColorFields`; `targetKey` and `lookup` let a field tell a new selection
 * from a re-render of its own drag, and put a drag it left back on its element.
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
	createInspectorColorFields,
	handleLiveColorEvent,
	inspectorFieldPatch,
	inspectorTogglePatch,
	linkedContentUrl,
	resolveInspectorSections,
} from 'ooxml-ui/pptx';
import type {
	InspectorFieldValue,
	InspectorSectionDescriptor,
	InspectorSectionView,
} from 'ooxml-ui/pptx';

import { isCheckboxControl, isSelectControl } from './control-event-targets';
import { injectLinkedContentAllowed } from './external-content.service';
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
								section.label ?? (section.labelKey ?? '' | translate)
							}}</span>
						</label>
					} @else if (section.label || section.labelKey) {
						<span class="pptx-ng-fields__section-title">{{
							section.label ?? (section.labelKey ?? '' | translate)
						}}</span>
					}
					@if (section.fields.length) {
						<div class="pptx-ng-fields__fields">
							@for (field of section.fields; track field.id) {
								@if (field.control === 'gallery') {
									<div
										class="pptx-ng-fields__gallery pptx-ng-fields__wide"
										role="group"
										[attr.aria-label]="field.labelKey | translate"
									>
										<span class="pptx-ng-fields__label">{{ field.labelKey | translate }}</span>
										<div class="pptx-ng-fields__gallery-grid" [attr.data-testid]="field.testId">
											@for (option of field.options ?? []; track option.value) {
												<button
													type="button"
													[class.is-active]="option.value === '' + field.value"
													[disabled]="disabled()"
													[attr.title]="option.label ?? (option.labelKey | translate)"
													[attr.aria-label]="option.label ?? (option.labelKey | translate)"
													[attr.aria-pressed]="option.value === '' + field.value"
													[attr.data-value]="option.value"
													(click)="onGallery(section.id, field.id, option.value)"
												>
													@if (previewUrl(field.previewImage); as preview) {
														<img
															[src]="preview"
															alt=""
															[style.filter]="option.previewFilter || null"
														/>
													} @else {
														<span
															class="pptx-ng-fields__swatch"
															[style.filter]="option.previewFilter || null"
														></span>
													}
													<span class="pptx-ng-fields__caption">{{
														option.label ?? (option.labelKey | translate)
													}}</span>
												</button>
											}
										</div>
									</div>
								} @else if (field.control === 'checkbox') {
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
														{{ option.label ?? (option.labelKey | translate) }}
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
										} @else if (field.control === 'textarea') {
											<textarea
												class="pptx-ng-fields__input pptx-ng-fields__input--text"
												rows="2"
												[attr.data-testid]="field.testId"
												[attr.aria-label]="field.labelKey | translate"
												[attr.placeholder]="
													field.placeholderKey ? (field.placeholderKey | translate) : null
												"
												[value]="field.value"
												[disabled]="disabled()"
												(input)="onField(section.id, field.id, $event)"
											></textarea>
										} @else if (field.control === 'text') {
											<input
												class="pptx-ng-fields__input pptx-ng-fields__input--text"
												type="text"
												[attr.data-testid]="field.testId"
												[attr.aria-label]="field.labelKey | translate"
												[attr.placeholder]="
													field.placeholderKey ? (field.placeholderKey | translate) : null
												"
												[value]="field.value"
												[disabled]="disabled()"
												(input)="onField(section.id, field.id, $event)"
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

		.pptx-ng-fields__input--text {
			box-sizing: border-box;
			width: 100%;
			min-width: 0;
			flex: 1 1 auto;
			font: inherit;
			resize: vertical;
		}

		.pptx-ng-fields__field:has(.pptx-ng-fields__input--text) {
			flex-direction: column;
			align-items: stretch;
		}

		.pptx-ng-fields__field:has(.pptx-ng-fields__input--text) .pptx-ng-fields__label {
			text-align: left;
		}

		.pptx-ng-fields__gallery {
			display: grid;
			gap: 3px;
		}

		.pptx-ng-fields__gallery-grid {
			display: grid;
			grid-template-columns: repeat(4, minmax(0, 1fr));
			gap: 4px;
		}

		.pptx-ng-fields__gallery-grid button {
			display: flex;
			flex-direction: column;
			align-items: center;
			gap: 2px;
			padding: 2px;
			border: 1px solid var(--pptx-inspector-border, #444);
			border-radius: 3px;
			background: none;
			color: inherit;
			cursor: pointer;
		}

		.pptx-ng-fields__gallery-grid button.is-active {
			border-color: var(--pptx-inspector-active, #0078d4);
		}

		.pptx-ng-fields__gallery-grid img,
		.pptx-ng-fields__swatch {
			width: 100%;
			height: 32px;
			object-fit: cover;
			border-radius: 2px;
			background: var(--pptx-inspector-input-bg, #2d2d2d);
		}

		.pptx-ng-fields__caption {
			width: 100%;
			overflow: hidden;
			font-size: 9px;
			text-align: center;
			text-overflow: ellipsis;
			white-space: nowrap;
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
	private readonly linkedContentAllowed = injectLinkedContentAllowed();
	/** A gallery preview picture; a linked one only once this viewer allows linked content. */
	protected previewUrl(url: string | undefined): string | undefined {
		return linkedContentUrl(url, this.linkedContentAllowed());
	}
	/** The identity of the element `target` belongs to (its id), for the colour fields. */
	readonly targetKey = input('');
	/** What the descriptors read on element `key` now: a drag the selection left goes back on it. */
	readonly lookup = input<(key: string) => T | undefined>(() => undefined);
	/**
	 * Emits a colour field's live patch while its picker is dragged, for the element `target`
	 * names (one the selection has left, when its drag is put back), WITHOUT an undo step.
	 * The drag's one `patch` follows when it settles, for the current target.
	 */
	readonly preview = output<{ patch: P; target: string }>();

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

	/** A gallery swatch was picked. */
	protected onGallery(sectionId: string, fieldId: string, value: string): void {
		this.emit(inspectorFieldPatch(this.sections(), sectionId, fieldId, this.target(), value));
	}

	/** One live colour controller per colour field (shared `createInspectorColorFields`). */
	private readonly colorField = createInspectorColorFields<T, P>({
		sections: () => this.sections(),
		target: () => this.target(),
		targetKey: () => this.targetKey(),
		lookup: (key) => this.lookup()(key),
		preview: (patch, target) => this.preview.emit({ patch, target }),
		commit: (patch) => this.patch.emit(patch),
		remember: (hex) => this.recentColors?.push(hex),
	});

	protected onColor(sectionId: string, fieldId: string, event: Event): void {
		handleLiveColorEvent(this.colorField(sectionId, fieldId), event);
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
	if (target instanceof HTMLTextAreaElement) {
		return target.value;
	}
	return isSelectControl(target) ? target.value : null;
}
