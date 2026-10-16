/**
 * chart-format-pane-host.component.ts: hosts the shared
 * `pptx-ui-chart-format-pane` (Format Chart Area / Data Series / Axis / ...)
 * while it is open and a chart is selected. The element owns the view and
 * hands back a whole new chart model per edit, applied here through the
 * editor's normal update + history path.
 */
import {
	ChangeDetectionStrategy,
	Component,
	computed,
	CUSTOM_ELEMENTS_SCHEMA,
	effect,
	inject,
	input,
} from '@angular/core';
import { TranslateService } from '@ngx-translate/core';
import type { ChartFormatChangeEvent, ChartFormatTargetEvent } from 'ooxml-ui/pptx';
import type { ChartPptxElement, PptxElement } from 'pptx-viewer-core';

import { ChartFormatPaneService } from './chart-format-pane.service';
import { EditorStateService } from './editor-state.service';
import { homeLanguage, ngxTranslate } from './ribbon-home-lang';

@Component({
	selector: 'pptx-chart-format-pane-host',
	standalone: true,
	changeDetection: ChangeDetectionStrategy.OnPush,
	host: { class: 'contents' },
	schemas: [CUSTOM_ELEMENTS_SCHEMA],
	template: `
		@if (chart(); as element) {
			<pptx-ui-chart-format-pane
				[chartData]="element.chartData"
				[target]="target()"
				[translateLabel]="translate()"
				[disabled]="!canEdit()"
				(chart-format-change)="onChange($event)"
				(chart-format-target)="onTarget($event)"
				(chart-format-close)="pane?.close()"
			/>
		}
	`,
})
export class ChartFormatPaneHostComponent {
	readonly element = input<PptxElement | null>(null);
	readonly slideIndex = input<number>(0);
	readonly canEdit = input<boolean>(false);
	/** False outside edit / master mode, where the pane never shows. */
	readonly active = input<boolean>(true);

	protected readonly pane = inject(ChartFormatPaneService, { optional: true });
	private readonly editor = inject(EditorStateService);
	private readonly translation = inject(TranslateService, { optional: true });
	private readonly language = homeLanguage(this.translation);

	private readonly selectedChart = computed<ChartPptxElement | null>(() => {
		const element = this.element();
		return element?.type === 'chart' && element.chartData ? element : null;
	});
	protected readonly chart = computed(() =>
		this.active() && this.pane?.open() ? this.selectedChart() : null,
	);
	protected readonly target = computed(() => this.pane?.targetFor(this.chart()?.id) ?? 'chartArea');
	protected readonly translate = computed(() => {
		this.language();
		return ngxTranslate(this.translation);
	});

	constructor() {
		// The pane closes as soon as the selection is no longer a chart.
		effect(() => {
			if (!this.selectedChart() && this.pane?.open()) {
				this.pane.close();
			}
		});
	}

	protected onChange(event: Event): void {
		const chart = this.chart();
		if (!chart || !this.canEdit()) {
			return;
		}
		const { chartData, target } = (event as ChartFormatChangeEvent).detail;
		this.pane?.setTarget(chart.id, target);
		this.editor.updateElement(this.slideIndex(), chart.id, { chartData } as Partial<PptxElement>);
	}

	protected onTarget(event: Event): void {
		const chart = this.chart();
		if (chart) {
			this.pane?.setTarget(chart.id, (event as ChartFormatTargetEvent).detail.target);
		}
	}
}
