/**
 * text-advanced-panel.component.ts: Standalone Angular component for editing
 * advanced text properties on a selected PPTX element: paragraph alignment, and
 * the shared Text Box options (`TEXT_BOX_SECTIONS`: vertical anchor, text
 * direction, autofit, wrap, columns) and character and paragraph spacing
 * (`CHARACTER_SPACING_SECTIONS`, `PARAGRAPH_SPACING_SECTIONS`: indentation,
 * spacing, line spacing, right-to-left) drawn by the generic
 * `pptx-inspector-fields` renderer.
 *
 * Selector: `pptx-text-advanced-panel`
 *
 * Ported from / models the patterns in:
 *   packages/react/src/viewer/components/inspector/TextPropertiesHelpers.tsx
 *   packages/react/src/viewer/components/inspector/TextProperties.tsx
 *   packages/angular/src/viewer/inspector-panel.component.ts
 *
 * Contract:
 *   [element]     : the selected PptxElement (required)
 *   (patch)       : emits a Partial<PptxElement> for the orchestrator to
 *                   commit via EditorStateService.updateElement
 *
 * @module viewer/text-advanced-panel
 */

import {
	ChangeDetectionStrategy,
	Component,
	computed,
	input,
	output,
	CUSTOM_ELEMENTS_SCHEMA,
} from '@angular/core';
import { TranslatePipe } from '@ngx-translate/core';
import {
	CHARACTER_SPACING_SECTIONS,
	PARAGRAPH_SPACING_SECTIONS,
	TEXT_BOX_SECTIONS,
	textStylePatch,
} from 'ooxml-ui/pptx';
import type { PptxElement, TextStyle } from 'pptx-viewer-core';
import { hasTextProperties } from 'pptx-viewer-core';

import { InspectorFieldsComponent } from './inspector-fields.component';
import { ALIGN_OPTIONS, alignPatch, textAdvancedStateOf } from './text-advanced-helpers';
import type { TextAdvancedState } from './text-advanced-helpers';

@Component({
	schemas: [CUSTOM_ELEMENTS_SCHEMA],
	selector: 'pptx-text-advanced-panel',
	standalone: true,
	changeDetection: ChangeDetectionStrategy.OnPush,
	imports: [TranslatePipe, InspectorFieldsComponent],
	template: `
		<div class="pptx-ng-txadv">
			@if (!hasText()) {
				<p class="pptx-ng-txadv__empty">{{ 'pptx.textAdvanced.selectPrompt' | translate }}</p>
			} @else {
				<!-- ── Alignment row ─────────────────────────────────────── -->
				<section class="pptx-ng-txadv__section">
					<h4 class="pptx-ng-txadv__heading">
						{{ 'pptx.textAdvanced.paragraphAlignment' | translate }}
					</h4>
					<div class="pptx-ng-txadv__row pptx-ng-txadv__row--wrap">
						@for (opt of alignOptions; track opt[0]) {
							<button
								type="button"
								class="pptx-ng-txadv__align-btn"
								[class.is-active]="state().align === opt[0]"
								[title]="opt[1]"
								(click)="onAlignChange(opt[0])"
							>
								{{ alignLabel(opt[0]) }}
							</button>
						}
					</div>
				</section>

				<!-- ── Text box options (shared descriptors) ───────────────── -->
				<section class="pptx-ng-txadv__section">
					<pptx-inspector-fields
						[sections]="textBoxSections"
						[target]="textStyle()"
						(patch)="onTextBoxPatch($event)"
					/>
				</section>

				<!-- ── Spacing, indentation and direction (shared descriptors) ── -->
				<section class="pptx-ng-txadv__section">
					<h4 class="pptx-ng-txadv__heading">{{ 'pptx.textAdvanced.spacing' | translate }}</h4>
					<pptx-inspector-fields
						[sections]="characterSpacingSections"
						[target]="textStyle()"
						(patch)="onTextBoxPatch($event)"
					/>
					<pptx-inspector-fields
						[sections]="paragraphSpacingSections"
						[target]="textStyle()"
						(patch)="onTextBoxPatch($event)"
					/>
				</section>
			}
		</div>
	`,
	styles: `
		.pptx-ng-txadv {
			display: flex;
			flex-direction: column;
			gap: 0;
			padding: 0.5rem;
			font-size: 12px;
			color: var(--pptx-inspector-fg, #e0e0e0);
		}

		.pptx-ng-txadv__empty {
			font-size: 11px;
			color: var(--pptx-inspector-muted, #888);
			margin: 0;
		}

		.pptx-ng-txadv__section {
			padding: 0.4rem 0;
			border-bottom: 1px solid var(--pptx-inspector-border, #333);
		}

		.pptx-ng-txadv__section:last-child {
			border-bottom: none;
		}

		.pptx-ng-txadv__heading {
			font-size: 10px;
			font-weight: 600;
			text-transform: uppercase;
			letter-spacing: 0.05em;
			color: var(--pptx-inspector-muted, #888);
			margin: 0 0 0.3rem 0;
		}

		.pptx-ng-txadv__row {
			display: flex;
			align-items: center;
			gap: 0.35rem;
			margin-bottom: 0.3rem;
		}

		.pptx-ng-txadv__row:last-child {
			margin-bottom: 0;
		}

		.pptx-ng-txadv__row--wrap {
			flex-wrap: wrap;
		}

		.pptx-ng-txadv__align-btn {
			height: 24px;
			min-width: 36px;
			padding: 0 6px;
			background: var(--pptx-inspector-input-bg, #2d2d2d);
			border: 1px solid var(--pptx-inspector-border, #444);
			color: inherit;
			border-radius: 3px;
			cursor: pointer;
			font-size: 10px;
			white-space: nowrap;
		}

		.pptx-ng-txadv__align-btn.is-active {
			background: var(--pptx-inspector-active, #0078d4);
			border-color: var(--pptx-inspector-active, #0078d4);
			color: #fff;
		}

		.pptx-ng-txadv__align-btn:hover:not(.is-active) {
			background: var(--pptx-inspector-hover, #3a3a3a);
		}

		/* ── Responsive: larger touch targets on coarse-pointer devices ── */
		@media (pointer: coarse), (max-width: 640px) {
			.pptx-ng-txadv {
				font-size: 14px;
			}

			.pptx-ng-txadv__align-btn {
				height: 36px;
				min-width: 48px;
				font-size: 12px;
			}
		}
	`,
})
export class TextAdvancedPanelComponent {
	/** The element whose advanced text properties are being edited. */
	readonly element = input.required<PptxElement>();

	/**
	 * Emits a Partial<PptxElement> patch each time the user commits a change.
	 * The orchestrator should call EditorStateService.updateElement(slideIndex, element().id, patch).
	 */
	readonly patch = output<Partial<PptxElement>>();

	/** Whether the element supports text properties. */
	protected readonly hasText = computed(() => hasTextProperties(this.element()));

	/** Derived advanced text state from the current element. */
	protected readonly state = computed<TextAdvancedState>(() => textAdvancedStateOf(this.element()));

	/** Exposed option arrays for the template. */
	protected readonly alignOptions = ALIGN_OPTIONS;
	protected readonly textBoxSections = TEXT_BOX_SECTIONS;
	protected readonly characterSpacingSections = CHARACTER_SPACING_SECTIONS;
	protected readonly paragraphSpacingSections = PARAGRAPH_SPACING_SECTIONS;

	/** The element's text style, the value the Text Box descriptors read. */
	protected readonly textStyle = computed<TextStyle | undefined>(() => {
		const el = this.element();
		return hasTextProperties(el) ? el.textStyle : undefined;
	});

	// ── Short display label for alignment buttons ──────────────────────────────

	protected alignLabel(align: NonNullable<TextStyle['align']>): string {
		switch (align) {
			case 'left':
				return '≡L';
			case 'center':
				return '≡C';
			case 'right':
				return '≡R';
			case 'justify':
				return '≡J';
			case 'justLow':
				return 'JL';
			case 'dist':
				return 'Di';
			case 'thaiDist':
				return 'TD';
			default:
				return '?';
		}
	}

	// ── Alignment ─────────────────────────────────────────────────────────────

	protected onAlignChange(align: NonNullable<TextStyle['align']>): void {
		this.emit(alignPatch(this.element(), align));
	}

	// ── Text box options ──────────────────────────────────────────────────────

	/** Merge a Text Box field's body-level text-style patch into the element. */
	protected onTextBoxPatch(patch: Partial<TextStyle>): void {
		this.emit(textStylePatch(this.element(), patch));
	}

	// ── Internal ──────────────────────────────────────────────────────────────

	private emit(p: Partial<PptxElement>): void {
		this.patch.emit(p);
	}
}
