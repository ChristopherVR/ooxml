/**
 * smart-art-properties.component.ts: SmartArt editing inspector (Angular).
 *
 * Selector: `pptx-smart-art-properties`
 *
 * Presentational panel for editing a selected SmartArt element's data model:
 * a layout switcher, colour-scheme select, style toggle (flat / moderate /
 * intense) and the text pane, which is the shared `pptx-ui-smartart-text-pane`
 * element (outline with levels, keyboard structure edits, move, sub-item,
 * remove, per-node style). Layout, colour and style edits go through the pure
 * helpers in `smart-art-properties-helpers.ts`; no editing logic lives here. Every edit
 * is emitted as a complete new `PptxSmartArtData` via `smartArtDataChange`, so
 * the parent (`InspectorPanelComponent`) can commit it through
 * `EditorStateService.updateElement()` as a single undo/redo history entry.
 *
 * Ported from the React inspector:
 *   packages/react/src/viewer/components/inspector/SmartArtPropertiesPanel.tsx
 *   packages/react/src/viewer/components/inspector/SmartArtLayoutSwitcher.tsx
 *
 * Usage:
 * ```html
 * <pptx-smart-art-properties
 *   [smartArtData]="data"
 *   [canEdit]="canEdit"
 *   (smartArtDataChange)="onSmartArtChange($event)"
 * />
 * ```
 *
 * @module angular-viewer/smart-art-properties
 */

import {
	ChangeDetectionStrategy,
	Component,
	computed,
	inject,
	input,
	output,
	CUSTOM_ELEMENTS_SCHEMA,
} from '@angular/core';
import { TranslatePipe, TranslateService } from '@ngx-translate/core';
import type { SmartArtTextPaneChangeEvent, SmartArtTextPaneViewState } from 'ooxml-ui/pptx';
import type {
	PptxSmartArtData,
	SmartArtColorScheme,
	SmartArtLayoutType,
	SmartArtStyle,
} from 'pptx-viewer-core';

import { isSelectControl } from './control-event-targets';
import { SWITCHABLE_LAYOUT_TYPES } from './editor-insert';
import { RecentColorsService } from './recent-colors.service';
import { homeLanguage, ngxTranslate } from './ribbon-home-lang';
import {
	smartArtColorSchemeLabelKey,
	smartArtLayoutLabelKey,
	smartArtStyleLabelKey,
} from './schema-token-labels';
import {
	currentColorScheme,
	currentLayout,
	currentStyle,
	setColorScheme,
	setLayout,
	setStyle,
	SMART_ART_COLOR_SCHEMES,
	SMART_ART_STYLE_OPTIONS,
} from './smart-art-properties-helpers';

@Component({
	schemas: [CUSTOM_ELEMENTS_SCHEMA],
	selector: 'pptx-smart-art-properties',
	standalone: true,
	changeDetection: ChangeDetectionStrategy.OnPush,
	imports: [TranslatePipe],
	template: `
		<section class="pptx-sa-props" [attr.aria-label]="'pptx.smartart.title' | translate">
			<h3 class="pptx-sa-props__heading">{{ 'pptx.smartart.title' | translate }}</h3>

			<!-- ── Layout switcher ──────────────────────────────────────────── -->
			<div class="pptx-sa-props__field">
				<span class="pptx-sa-props__label">{{ 'pptx.master.layout' | translate }}</span>
				<div
					class="pptx-sa-props__layouts"
					role="group"
					[attr.aria-label]="'pptx.smartart.switchLayout' | translate"
				>
					@for (layout of layoutTypes; track layout) {
						<button
							type="button"
							[attr.data-testid]="'smartart-layout-' + layout"
							class="pptx-sa-props__layout"
							[class.is-active]="activeLayout() === layout"
							[disabled]="!canEdit()"
							[attr.aria-pressed]="activeLayout() === layout"
							(click)="onLayout(layout)"
						>
							{{ layoutLabelKey(layout) | translate }}
						</button>
					}
				</div>
			</div>

			<!-- ── Colour scheme ────────────────────────────────────────────── -->
			<label class="pptx-sa-props__field">
				<span class="pptx-sa-props__label">{{ 'pptx.smartart.colorScheme' | translate }}</span>
				<pptx-ui-select
					class="pptx-sa-props__select"
					data-testid="smartart-color-scheme"
					[attr.aria-label]="'pptx.smartart.colorScheme' | translate"
					[disabled]="!canEdit()"
					[value]="activeColorScheme()"
					(change)="onColorScheme($event)"
				>
					@for (scheme of colorSchemes; track scheme) {
						<option [value]="scheme" [selected]="scheme === activeColorScheme()">
							{{ colorSchemeLabelKey(scheme) | translate }}
						</option>
					}
				</pptx-ui-select>
			</label>

			<!-- ── Style intensity ──────────────────────────────────────────── -->
			<div class="pptx-sa-props__field">
				<span class="pptx-sa-props__label">{{ 'pptx.smartart.style' | translate }}</span>
				<div
					class="pptx-sa-props__styles"
					role="group"
					[attr.aria-label]="'pptx.smartart.style' | translate"
				>
					@for (styleOpt of styleOptions; track styleOpt) {
						<button
							type="button"
							class="pptx-sa-props__style"
							[class.is-active]="activeStyle() === styleOpt"
							[disabled]="!canEdit()"
							[attr.aria-pressed]="activeStyle() === styleOpt"
							(click)="onStyle(styleOpt)"
						>
							{{ styleLabelKey(styleOpt) | translate }}
						</button>
					}
				</div>
			</div>

			<!-- ── Text pane: the shared element ─────────────────────────────── -->
			<pptx-ui-smartart-text-pane
				data-testid="smartart-text-pane"
				[state]="paneState()"
				(pptx-smartart-data-change)="onPaneChange($event)"
			></pptx-ui-smartart-text-pane>
		</section>
	`,
	styles: `
		.pptx-sa-props {
			display: flex;
			flex-direction: column;
			gap: 0.4rem;
			padding: 0.5rem 0;
		}

		.pptx-sa-props__heading {
			font-size: 10px;
			font-weight: 600;
			text-transform: uppercase;
			letter-spacing: 0.05em;
			color: var(--pptx-inspector-muted, #888);
			margin: 0;
		}

		.pptx-sa-props__field {
			display: flex;
			flex-direction: column;
			gap: 0.2rem;
		}

		.pptx-sa-props__label {
			font-size: 10px;
			color: var(--pptx-inspector-muted, #888);
		}

		.pptx-sa-props__layouts,
		.pptx-sa-props__styles {
			display: flex;
			flex-wrap: wrap;
			gap: 0.25rem;
		}

		.pptx-sa-props__layout,
		.pptx-sa-props__style {
			background: var(--pptx-inspector-input-bg, #2d2d2d);
			border: 1px solid var(--pptx-inspector-border, #444);
			color: inherit;
			border-radius: 3px;
			cursor: pointer;
			font-size: 10px;
			padding: 2px 6px;
			white-space: nowrap;
		}

		/*
		 * No text-transform: these buttons used to print the raw wire token
		 * ("list", "flat"), and capitalising it was the only thing that made it
		 * look like a word. They now render dictionary text, which arrives cased
		 * for its locale, so capitalising would mangle multi-word translations.
		 */
		.pptx-sa-props__layout,
		.pptx-sa-props__style {
			flex: 1 0 auto;
		}

		.pptx-sa-props__layout.is-active,
		.pptx-sa-props__style.is-active {
			background: var(--pptx-inspector-active, #0078d4);
			border-color: var(--pptx-inspector-active, #0078d4);
			color: #fff;
		}

		.pptx-sa-props__select {
			background: var(--pptx-inspector-input-bg, #2d2d2d);
			border: 1px solid var(--pptx-inspector-border, #444);
			color: inherit;
			border-radius: 3px;
			padding: 2px 4px;
			font-size: 11px;
		}

		.pptx-sa-props__layout:disabled,
		.pptx-sa-props__style:disabled,
		.pptx-sa-props__select:disabled {
			opacity: 0.45;
			cursor: not-allowed;
		}
	`,
})
export class SmartArtPropertiesComponent {
	/** The SmartArt data model being edited. */
	readonly smartArtData = input.required<PptxSmartArtData>();
	/** Whether editing is enabled (read-only when false). */
	readonly canEdit = input<boolean>(true);

	/** Emits a complete new data model after any edit. */
	readonly smartArtDataChange = output<PptxSmartArtData>();

	/** Optional: absent in a standalone unit test with no viewer-level DI tree. */
	private readonly recentColors = inject(RecentColorsService, { optional: true });
	private readonly translate = inject(TranslateService, { optional: true });
	/** Read so a language switch re-labels the text pane. */
	private readonly language = homeLanguage(this.translate);

	// ── Static option lists (template-bound) ─────────────────────────────────
	protected readonly layoutTypes: readonly SmartArtLayoutType[] = SWITCHABLE_LAYOUT_TYPES;
	protected readonly colorSchemes = SMART_ART_COLOR_SCHEMES;
	protected readonly styleOptions = SMART_ART_STYLE_OPTIONS;

	// ── Derived state ─────────────────────────────────────────────────────────
	protected readonly activeLayout = computed(() => currentLayout(this.smartArtData()));
	protected readonly activeColorScheme = computed(() => currentColorScheme(this.smartArtData()));
	protected readonly activeStyle = computed(() => currentStyle(this.smartArtData()));

	/** What the shared text pane shows. */
	protected readonly paneState = computed<SmartArtTextPaneViewState>(() => {
		this.language();
		return {
			smartArtData: this.smartArtData(),
			canEdit: this.canEdit(),
			translate: ngxTranslate(this.translate),
		};
	});

	// ── Wire-token spelling ──────────────────────────────────────────────────
	// These three controls used to print the `dgm:` schema token itself, so the
	// picker offered "colorful1" and "bending" as if they were English words.
	protected layoutLabelKey = smartArtLayoutLabelKey;
	protected colorSchemeLabelKey = smartArtColorSchemeLabelKey;
	protected styleLabelKey = smartArtStyleLabelKey;

	// ── Event handlers (each emits a fresh data model) ───────────────────────
	protected onLayout(layout: SmartArtLayoutType): void {
		this.commit(setLayout(this.smartArtData(), layout));
	}

	protected onColorScheme(event: Event): void {
		const value = selectValue(event);
		if (value === null) {
			return;
		}
		this.commit(setColorScheme(this.smartArtData(), value as SmartArtColorScheme));
	}

	protected onStyle(style: SmartArtStyle): void {
		this.commit(setStyle(this.smartArtData(), style));
	}

	/** A text pane edit: the element already applied it through the pptx SmartArt facades. */
	protected onPaneChange(event: Event): void {
		const { smartArtData, recentColor } = (event as SmartArtTextPaneChangeEvent).detail;
		this.commit(smartArtData);
		if (recentColor) {
			this.recentColors?.push(recentColor);
		}
	}

	/** Emit a new data model only when it differs from the current one. */
	private commit(next: PptxSmartArtData): void {
		if (!this.canEdit() || next === this.smartArtData()) {
			return;
		}
		this.smartArtDataChange.emit(next);
	}
}

// ── Module-private helpers ───────────────────────────────────────────────────

/** Read the value of a `<pptx-ui-select>` change event. */
function selectValue(event: Event): string | null {
	const target = event.target;
	return isSelectControl(target) ? target.value : null;
}
