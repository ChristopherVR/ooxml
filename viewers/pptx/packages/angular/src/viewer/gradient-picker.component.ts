/**
 * gradient-picker.component.ts: Standalone Angular component for a shape's
 * gradient stop list: one row per stop (colour with its theme grid,
 * position, opacity, remove) and an Add Stop button.
 *
 * Selector: `pptx-gradient-picker`
 *
 * The rows are the shared `resolveShapeFillView` stop views and every control
 * is a shared `shapeFillPatch` action, merged into the element with
 * `shapeQuickStylePatch`. The gradient type and angle are shared fill
 * descriptor fields, drawn above it by `pptx-shape-fill-section`.
 *
 * Contract:
 *   [element]     : the selected PptxElement (required)
 *   (patch)       : emits a Partial<PptxElement> for the orchestrator to
 *                   commit via EditorStateService.updateElement
 *
 * @module viewer/gradient-picker
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
import { TranslatePipe } from '@ngx-translate/core';
import {
	resolveShapeFillView,
	shapeFillModeChange,
	shapeFillPatch,
	shapeQuickStylePatch,
} from 'ooxml-ui/pptx';
import type { ShapeFillAction, ThemeColorPickerCommit } from 'ooxml-ui/pptx';
import type { PptxElement } from 'pptx-viewer-core';
import { hasShapeProperties } from 'pptx-viewer-core';

import { RecentColorsService } from './recent-colors.service';
import { ThemeColorSwatchGridComponent } from './theme-color-swatch-grid.component';

@Component({
	schemas: [CUSTOM_ELEMENTS_SCHEMA],
	selector: 'pptx-gradient-picker',
	standalone: true,
	imports: [TranslatePipe, ThemeColorSwatchGridComponent],
	changeDetection: ChangeDetectionStrategy.OnPush,
	template: `
		<div class="pptx-ng-grad">
			<div class="pptx-ng-grad__stops-heading">
				{{ 'pptx.fillAdvanced.gradientStops' | translate }}
			</div>
			@for (stop of stops(); track stop.index) {
				<div class="pptx-ng-grad__stop-block" data-testid="fill-gradient-stop">
					<div class="pptx-ng-grad__stop-row">
						<span class="pptx-ng-grad__stop-idx">{{ stop.index + 1 }}</span>
						<input
							class="pptx-ng-grad__color"
							type="color"
							[attr.aria-label]="'pptx.gradient.stops' | translate"
							[value]="stop.color"
							(change)="onStopColor($event, stop.index)"
						/>
						<input
							class="pptx-ng-grad__input pptx-ng-grad__input--number"
							type="number"
							inputmode="numeric"
							min="0"
							max="100"
							step="1"
							[attr.aria-label]="'pptx.gradient.position' | translate"
							[value]="stop.position"
							(change)="onPosition($event, stop.index)"
						/>
						<input
							class="pptx-ng-grad__input pptx-ng-grad__input--number"
							type="number"
							inputmode="numeric"
							min="0"
							max="100"
							step="1"
							[attr.aria-label]="'pptx.strokeEffects.fillOpacity' | translate"
							[value]="stop.opacity"
							(change)="onOpacity($event, stop.index)"
						/>
						<button
							type="button"
							class="pptx-ng-grad__btn pptx-ng-grad__btn--remove"
							data-pptx-compact
							[disabled]="!stop.removable"
							[title]="'pptx.gradient.removeStop' | translate"
							[attr.aria-label]="'pptx.gradient.removeStop' | translate"
							(click)="onRemove(stop.index)"
						>
							×
						</button>
					</div>
					<pptx-theme-color-swatch-grid
						[selectedRef]="stop.colorRef"
						[selectedHex]="stop.color"
						(pick)="onStopThemePick($event, stop.index)"
					/>
				</div>
			}
			<button type="button" class="pptx-ng-grad__btn pptx-ng-grad__btn--add" (click)="onAdd()">
				{{ 'pptx.gradient.addStop' | translate }}
			</button>
		</div>
	`,
	styles: `
		.pptx-ng-grad {
			display: flex;
			flex-direction: column;
			gap: 0.35rem;
			padding: 0.5rem;
			font-size: 12px;
			color: var(--pptx-inspector-fg, #e0e0e0);
		}

		.pptx-ng-grad__input {
			background: var(--pptx-inspector-input-bg, #2d2d2d);
			border: 1px solid var(--pptx-inspector-border, #444);
			color: inherit;
			border-radius: 3px;
			padding: 2px 4px;
			font-size: 12px;
		}

		.pptx-ng-grad__input--number {
			width: 52px;
			text-align: right;
		}

		.pptx-ng-grad__stops-heading {
			font-size: 10px;
			text-transform: uppercase;
			letter-spacing: 0.04em;
			color: var(--pptx-inspector-muted, #888);
			margin-top: 0.25rem;
		}

		.pptx-ng-grad__stop-block {
			display: flex;
			flex-direction: column;
			gap: 0.2rem;
		}

		.pptx-ng-grad__stop-row {
			display: flex;
			align-items: center;
			gap: 0.25rem;
		}

		.pptx-ng-grad__stop-idx {
			font-size: 10px;
			color: var(--pptx-inspector-muted, #888);
			min-width: 14px;
			text-align: center;
			flex-shrink: 0;
		}

		.pptx-ng-grad__color {
			width: 28px;
			height: 22px;
			border: 1px solid var(--pptx-inspector-border, #444);
			border-radius: 3px;
			padding: 1px;
			cursor: pointer;
			background: transparent;
			flex-shrink: 0;
		}

		.pptx-ng-grad__btn {
			background: var(--pptx-inspector-input-bg, #2d2d2d);
			border: 1px solid var(--pptx-inspector-border, #444);
			color: inherit;
			border-radius: 3px;
			cursor: pointer;
			font-size: 11px;
			padding: 2px 6px;
		}

		.pptx-ng-grad__btn--add {
			align-self: flex-start;
			margin-top: 0.25rem;
		}

		.pptx-ng-grad__btn--remove {
			width: 20px;
			height: 20px;
			padding: 0;
			text-align: center;
			color: var(--pptx-inspector-danger, #f47c7c);
			border-color: var(--pptx-inspector-danger-border, #6b2a2a);
			flex-shrink: 0;
		}

		.pptx-ng-grad__btn:hover {
			background: var(--pptx-inspector-hover, #3a3a3a);
		}

		.pptx-ng-grad__btn--remove:hover {
			background: var(--pptx-inspector-danger-hover, #4a1a1a);
		}
	`,
})
export class GradientPickerComponent {
	/** The element whose gradient fill is being edited. */
	readonly element = input.required<PptxElement>();

	/**
	 * Emits a Partial<PptxElement> patch each time the user commits a change.
	 * The orchestrator (InspectorPanelComponent or a wrapper) should call
	 * EditorStateService.updateElement(slideIndex, element().id, patch).
	 */
	readonly patch = output<Partial<PptxElement>>();

	/** Optional: absent in a standalone unit test with no viewer-level DI tree. */
	private readonly recentColors = inject(RecentColorsService, { optional: true });

	private readonly style = computed(() => {
		const el = this.element();
		return hasShapeProperties(el) ? el.shapeStyle : undefined;
	});

	/** The shared stop rows; a shape not yet in a gradient fill shows the stops it would get. */
	protected readonly stops = computed(() => {
		const style = this.style();
		return (
			resolveShapeFillView(style).stops ??
			resolveShapeFillView({ ...style, ...shapeFillModeChange(style, 'gradient') }).stops ??
			[]
		);
	});

	/** Apply a shared fill action to the element. */
	private act(action: ShapeFillAction): void {
		const patch = shapeFillPatch(this.style(), action);
		if (patch) {
			this.patch.emit(shapeQuickStylePatch(this.element(), patch));
		}
	}

	/** A native pick has no theme identity, so the stop's stored `colorRef` is cleared. */
	protected onStopColor(event: Event, index: number): void {
		const target = event.target;
		if (!(target instanceof HTMLInputElement) || !target.value) {
			return;
		}
		this.act({ type: 'stopColor', index, hex: target.value });
		this.recentColors?.push(target.value);
	}

	/** A theme-swatch pick: commits BOTH the resolved hex and the ref. */
	protected onStopThemePick(commit: ThemeColorPickerCommit, index: number): void {
		this.act({ type: 'stopColor', index, hex: commit.hex, ref: commit.ref });
		this.recentColors?.push(commit.hex);
	}

	protected onPosition(event: Event, index: number): void {
		this.act({ type: 'stopPosition', index, position: numberOf(event) });
	}

	protected onOpacity(event: Event, index: number): void {
		this.act({ type: 'stopOpacity', index, opacity: numberOf(event) });
	}

	protected onRemove(index: number): void {
		this.act({ type: 'removeStop', index });
	}

	protected onAdd(): void {
		this.act({ type: 'addStop' });
	}
}

/** The number a box holds (`NaN` when empty, which the shared action ignores). */
function numberOf(event: Event): number {
	const target = event.target;
	return target instanceof HTMLInputElement && target.value.trim() !== ''
		? Number(target.value)
		: Number.NaN;
}
