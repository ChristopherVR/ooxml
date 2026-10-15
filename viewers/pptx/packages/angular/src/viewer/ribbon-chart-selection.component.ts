/**
 * ribbon-chart-selection.component.ts: Chart Format > Current Selection, a
 * thin adapter around `pptx-ui-ribbon-home-chart-selection` (the element
 * drop-down, Format Selection and Reset to Match Style). The decisions live
 * in the shared `chartSelectionHomeControls` / `dispatchChartSelectionRequest`.
 */
import {
	ChangeDetectionStrategy,
	Component,
	CUSTOM_ELEMENTS_SCHEMA,
	inject,
	input,
} from '@angular/core';
import { TranslateService } from '@ngx-translate/core';
import {
	chartSelectionHomeControls,
	dispatchChartSelectionRequest,
	resetChartFormatTarget,
} from 'ooxml-ui/pptx';
import type { RibbonHomeRequestEvent } from 'ooxml-ui/pptx';
import type { PptxElement } from 'pptx-viewer-core';

import { ChartFormatPaneService } from './chart-format-pane.service';
import { EditorStateService } from './editor-state.service';
import { homeLanguage, homeTranslator, ngxTranslate } from './ribbon-home-lang';

@Component({
	selector: 'pptx-ribbon-chart-selection',
	standalone: true,
	changeDetection: ChangeDetectionStrategy.OnPush,
	host: { class: 'contents' },
	schemas: [CUSTOM_ELEMENTS_SCHEMA],
	template: `<pptx-ui-ribbon-home-chart-selection
		[state]="view()"
		(home-request)="request($event)"
	/>`,
})
export class RibbonChartSelectionComponent {
	private readonly translation = inject(TranslateService, { optional: true });
	private readonly language = homeLanguage(this.translation);
	private readonly editor = inject(EditorStateService);
	private readonly pane = inject(ChartFormatPaneService, { optional: true });

	readonly slideIndex = input<number>(0);
	readonly selectedElement = input<PptxElement | null>(null);
	readonly canEdit = input<boolean>(false);

	private chart() {
		const element = this.selectedElement();
		return element?.type === 'chart' && element.chartData ? element : null;
	}

	protected view() {
		const chart = this.chart();
		this.language();
		return {
			controls: chartSelectionHomeControls({
				chartData: chart?.chartData,
				targetKey: this.pane?.targetFor(chart?.id),
				editable: this.canEdit(),
				translate: ngxTranslate(this.translation),
			}),
			translate: homeTranslator(this.translation, this.language, ['chart-selection']),
		};
	}

	protected request(event: Event): void {
		const chart = this.chart();
		if (!chart?.chartData) {
			return;
		}
		dispatchChartSelectionRequest((event as RibbonHomeRequestEvent).detail, {
			select: (target) => this.pane?.setTarget(chart.id, target),
			formatSelection: () => this.pane?.openOn(chart.id, this.pane.targetFor(chart.id)),
			reset: () => {
				if (!this.canEdit() || !chart.chartData) {
					return;
				}
				const chartData = resetChartFormatTarget(
					chart.chartData,
					this.pane?.targetFor(chart.id) ?? 'chartArea',
				);
				this.editor.updateElement(this.slideIndex(), chart.id, {
					chartData,
				} as Partial<PptxElement>);
			},
		});
	}
}
