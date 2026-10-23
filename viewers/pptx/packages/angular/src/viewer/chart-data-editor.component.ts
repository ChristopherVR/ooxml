/**
 * chart-data-editor.component.ts: the chart's "Edit Data" panel.
 *
 * Selector: `pptx-chart-data-editor`
 *
 * The categories x series sheet is the shared `pptx-ui-chart-data-sheet` element (keyboard
 * navigation, tab-separated paste, locale numbers), the same one every binding shows; this
 * component feeds it the chart and the canvas part selection, and emits the edited chart as a
 * complete new `ChartPptxElement` via `elementChange`. Below the sheet it keeps the per-series
 * colour pickers, the Chart Filters options and the advanced chart editor.
 *
 * The parent (typically `InspectorPanelComponent`) receives the emitted element and commits it
 * to `EditorStateService.updateElement()` as a single history entry.
 *
 * Usage:
 * ```html
 * <pptx-chart-data-editor
 *   [element]="selectedElement"
 *   [canEdit]="canEdit"
 *   (elementChange)="onChartChange($event)"
 * />
 * ```
 *
 * @module angular-viewer/chart-data-editor
 */

import {
	ChangeDetectionStrategy,
	Component,
	CUSTOM_ELEMENTS_SCHEMA,
	computed,
	inject,
	input,
	output,
} from '@angular/core';
import { TranslatePipe, TranslateService } from '@ngx-translate/core';
import { seriesColorClearState } from 'ooxml-ui/pptx';
import type {
	ChartDataSheetChangeEvent,
	ChartDataSheetViewState,
	InspectorActionState,
} from 'ooxml-ui/pptx';
import type { ChartPptxElement, PptxChartSeries } from 'pptx-viewer-core';

import { AdvancedChartEditorComponent } from './advanced-chart-editor.component';
import { setSeriesColor } from './chart-data-helpers';
import { ChartFilteredSeriesOptionsComponent } from './chart-filtered-series-options.component';
import { ChartPartSelectionService } from './chart-part-selection.service';
import { RecentColorsService } from './recent-colors.service';
import { homeLanguage, ngxTranslate } from './ribbon-home-lang';
import { ViewerOptionsService } from './viewer-options.service';

@Component({
	selector: 'pptx-chart-data-editor',
	standalone: true,
	imports: [AdvancedChartEditorComponent, ChartFilteredSeriesOptionsComponent, TranslatePipe],
	changeDetection: ChangeDetectionStrategy.OnPush,
	schemas: [CUSTOM_ELEMENTS_SCHEMA],
	template: `
		<section class="pptx-chart-editor" [attr.aria-label]="'pptx.chart.data' | translate">
			@if (hasData()) {
				<pptx-ui-chart-data-sheet
					[state]="sheetState()"
					(pptx-chart-data-change)="onSheetChange($event)"
				></pptx-ui-chart-data-sheet>
				<div class="pptx-chart-editor__colors">
					@for (s of series(); track $index; let si = $index) {
						<span class="pptx-chart-editor__color-wrap">
							<input
								type="color"
								class="pptx-chart-editor__color-input"
								[disabled]="!canEdit()"
								[value]="s.color || '#4472c4'"
								[title]="'pptx.chart.seriesColor' | translate: { name: seriesDisplayName(s, si) }"
								[attr.aria-label]="
									'pptx.chart.seriesColor' | translate: { name: seriesDisplayName(s, si) }
								"
								(input)="onSeriesColorChange($event, si)"
								(change)="pushRecentColor($event)"
							/>
							@if (seriesClearState(s).visible) {
								<button
									type="button"
									class="pptx-chart-editor__remove-btn"
									[title]="'pptx.chart.clearSeriesColor' | translate"
									[attr.aria-label]="'pptx.chart.clearSeriesColor' | translate"
									(click)="onClearSeriesColor(si)"
								>
									×
								</button>
							}
						</span>
					}
				</div>
				<pptx-chart-filtered-series-options
					[element]="element()"
					[canEdit]="canEdit()"
					(elementChange)="elementChange.emit($event)"
				/>
				<pptx-advanced-chart-editor
					[element]="element()"
					[canEdit]="canEdit()"
					(elementChange)="elementChange.emit($event)"
				/>
			} @else {
				<p class="pptx-chart-editor__empty">{{ 'pptx.chart.noData' | translate }}</p>
			}
		</section>
	`,
	styles: `
		.pptx-chart-editor {
			display: flex;
			flex-direction: column;
			gap: 0.35rem;
			padding: 0.5rem 0;
			border-bottom: 1px solid var(--pptx-inspector-border, #333);
		}

		.pptx-chart-editor__colors {
			display: flex;
			flex-wrap: wrap;
			gap: 0.25rem;
		}

		.pptx-chart-editor__color-wrap {
			display: inline-flex;
			align-items: center;
			gap: 1px;
		}

		.pptx-chart-editor__color-input {
			width: 22px;
			height: 18px;
			padding: 0;
			border: 1px solid var(--pptx-inspector-border, #444);
			border-radius: 3px;
			background: transparent;
			cursor: pointer;
		}

		.pptx-chart-editor__color-input:disabled {
			opacity: 0.6;
			cursor: not-allowed;
		}

		.pptx-chart-editor__remove-btn {
			padding: 0 2px;
			font-size: 11px;
			line-height: 1;
			background: none;
			border: none;
			color: var(--pptx-inspector-danger, #f47c7c);
			cursor: pointer;
		}

		.pptx-chart-editor__empty {
			font-size: 11px;
			color: var(--pptx-inspector-muted, #888);
			margin: 0.25rem 0;
		}
	`,
})
export class ChartDataEditorComponent {
	private readonly translate = inject(TranslateService);
	/** Read so a language switch re-labels the sheet. */
	private readonly language = homeLanguage(this.translate);
	/** Canvas <-> inspector chart-part selection bridge (viewer-scoped, optional). */
	private readonly partSelection = inject(ChartPartSelectionService, { optional: true });
	/** File > Options access (viewer-scoped, optional for standalone unit tests). */
	private readonly viewerOpts = inject(ViewerOptionsService, { optional: true });
	/** Optional: absent in a standalone unit test with no viewer-level DI tree. */
	private readonly recentColors = inject(RecentColorsService, { optional: true });

	/** The chart element being edited. */
	readonly element = input.required<ChartPptxElement>();
	/** Whether editing is enabled (read-only mode when false). */
	readonly canEdit = input<boolean>(true);

	/** Emits the updated element after any edit operation. */
	readonly elementChange = output<ChartPptxElement>();

	/**
	 * Cell to highlight, driven by the on-canvas chart part selection: a `pointIndex` highlights
	 * one value cell, series-only highlights the series name cell.
	 */
	protected readonly highlightCell = computed<{ seriesIndex: number; pointIndex?: number } | null>(
		() => {
			const sel = this.partSelection?.selection() ?? null;
			if (!sel || sel.elementId !== this.element().id) {
				return null;
			}
			return { seriesIndex: sel.part.seriesIndex, pointIndex: sel.part.pointIndex };
		},
	);

	protected readonly sheetState = computed<ChartDataSheetViewState>(() => {
		this.language();
		return {
			chartData: this.element().chartData,
			canEdit: this.canEdit(),
			// File > Options > Advanced > "Properties follow chart data point for current
			// workbook"; PowerPoint's own default (`true`) without the service (unit tests).
			followDataPoint: this.viewerOpts?.options().advanced.chartPropertiesFollowDataPoint ?? true,
			highlight: this.highlightCell(),
			translate: ngxTranslate(this.translate),
		};
	});

	protected readonly hasData = computed(() => {
		const data = this.element().chartData;
		return data !== undefined && data.series.length > 0 && data.categories.length > 0;
	});

	protected readonly series = computed(() => this.element().chartData?.series ?? []);

	/** Series display name, falling back to a translated "Series N" when unnamed. */
	protected seriesDisplayName(series: PptxChartSeries, index: number): string {
		return (
			series.name || this.translate.instant('pptx.chart.seriesDefaultName', { number: index + 1 })
		);
	}

	protected onSheetChange(event: Event): void {
		const chartData = (event as ChartDataSheetChangeEvent).detail.chartData;
		this.elementChange.emit({ ...this.element(), chartData });
	}

	protected onSeriesColorChange(event: Event, seriesIndex: number): void {
		const target = event.target;
		if (target instanceof HTMLInputElement) {
			this.elementChange.emit(setSeriesColor(this.element(), seriesIndex, target.value));
		}
	}

	protected seriesClearState(series: PptxChartSeries): InspectorActionState {
		return seriesColorClearState(series, this.canEdit());
	}

	protected onClearSeriesColor(seriesIndex: number): void {
		this.elementChange.emit(setSeriesColor(this.element(), seriesIndex, null));
	}

	/**
	 * Record the committed (native `change`, not the live-preview `input`) series colour into the
	 * shared "Recent colours" list.
	 */
	protected pushRecentColor(event: Event): void {
		const target = event.target;
		if (target instanceof HTMLInputElement && target.value) {
			this.recentColors?.push(target.value);
		}
	}
}
