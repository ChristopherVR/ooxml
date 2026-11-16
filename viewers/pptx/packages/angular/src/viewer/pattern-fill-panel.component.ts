/**
 * pattern-fill-panel.component.ts: Standalone Angular component for editing a
 * shape's pattern fill (`a:pattFill`): preset swatches plus foreground /
 * background colours.
 *
 * Selector: `pptx-pattern-fill-panel`
 *
 * The swatches (56 presets, each tile in the shape's own colours) and colours
 * are the shared `resolveShapeFillView` pattern view, and every control is a
 * shared `shapeFillPatch` action merged into the element with
 * `shapeQuickStylePatch`.
 *
 * Contract:
 *   [element]     : the selected PptxElement (required)
 *   (patch)       : emits a Partial<PptxElement> for the orchestrator to
 *                   commit via EditorStateService.updateElement
 *
 * @module viewer/pattern-fill-panel
 */
import { ChangeDetectionStrategy, Component, computed, inject, input, output } from '@angular/core';
import { TranslatePipe } from '@ngx-translate/core';
import {
	resolveShapeFillView,
	shapeFillModeChange,
	shapeFillPatch,
	shapeQuickStylePatch,
} from 'ooxml-ui/pptx';
import type { ShapeFillAction, ShapeFillView } from 'ooxml-ui/pptx';
import type { PptxElement, ShapeStyle } from 'pptx-viewer-core';
import { hasShapeProperties } from 'pptx-viewer-core';

import { RecentColorsService } from './recent-colors.service';

@Component({
	selector: 'pptx-pattern-fill-panel',
	standalone: true,
	imports: [TranslatePipe],
	changeDetection: ChangeDetectionStrategy.OnPush,
	template: `
		<div class="pptx-ng-pf">
			@if (!applicable()) {
				<p class="pptx-ng-pf__muted">{{ 'pptx.fill.noOptions' | translate }}</p>
			} @else {
				<span class="pptx-ng-pf__label">{{ 'pptx.table.patternPreset' | translate }}</span>
				<div class="pptx-ng-pf__grid">
					@for (swatch of pattern().swatches; track swatch.value) {
						<button
							type="button"
							data-testid="fill-pattern-swatch"
							class="pptx-ng-pf__swatch"
							[class.is-active]="swatch.selected"
							[title]="swatch.labelKey | translate"
							[attr.aria-label]="swatch.labelKey | translate"
							[attr.aria-pressed]="swatch.selected"
							(click)="onPreset(swatch.value)"
						>
							@if (swatch.backgroundImage; as url) {
								<span class="pptx-ng-pf__swatch-fill" [style.background-image]="url"></span>
							}
						</button>
					}
				</div>

				<label class="pptx-ng-pf__field" for="pf-fg">
					<span class="pptx-ng-pf__label">{{
						'pptx.fillAdvanced.foregroundColor' | translate
					}}</span>
					<input
						id="pf-fg"
						type="color"
						class="pptx-ng-pf__color"
						[value]="pattern().foreground"
						(change)="onColor($event, 'patternForeground')"
					/>
				</label>
				<label class="pptx-ng-pf__field" for="pf-bg">
					<span class="pptx-ng-pf__label">{{
						'pptx.fillAdvanced.backgroundColor' | translate
					}}</span>
					<input
						id="pf-bg"
						type="color"
						class="pptx-ng-pf__color"
						[value]="pattern().background"
						(change)="onColor($event, 'patternBackground')"
					/>
				</label>
			}
		</div>
	`,
	styles: `
		.pptx-ng-pf {
			display: flex;
			flex-direction: column;
			gap: 0.4rem;
			padding: 0.5rem;
			font-size: 12px;
			color: var(--pptx-inspector-fg, #e0e0e0);
		}
		.pptx-ng-pf__muted {
			color: var(--pptx-inspector-muted, #888);
			font-style: italic;
			margin: 0;
		}
		.pptx-ng-pf__label {
			font-size: 10px;
			color: var(--pptx-inspector-muted, #888);
		}
		.pptx-ng-pf__grid {
			display: grid;
			grid-template-columns: repeat(8, 1fr);
			gap: 3px;
			max-height: 180px;
			overflow-y: auto;
			border: 1px solid var(--pptx-inspector-border, #444);
			border-radius: 3px;
			padding: 4px;
		}
		.pptx-ng-pf__swatch {
			width: 20px;
			height: 20px;
			border: 1px solid var(--pptx-inspector-border, #444);
			border-radius: 2px;
			background: #fff;
			padding: 0;
			cursor: pointer;
		}
		.pptx-ng-pf__swatch.is-active {
			border: 2px solid var(--pptx-inspector-accent, #2f6feb);
		}
		.pptx-ng-pf__swatch-fill {
			display: block;
			width: 100%;
			height: 100%;
			background-repeat: repeat;
			background-size: 8px 8px;
		}
		.pptx-ng-pf__field {
			display: flex;
			flex-direction: column;
			gap: 0.2rem;
		}
		.pptx-ng-pf__color {
			width: 100%;
			height: 24px;
			border: 1px solid var(--pptx-inspector-border, #444);
			border-radius: 3px;
			padding: 1px;
			cursor: pointer;
			background: transparent;
		}
	`,
})
export class PatternFillPanelComponent {
	readonly element = input.required<PptxElement>();
	readonly patch = output<Partial<PptxElement>>();

	/** Optional: absent in a standalone unit test with no viewer-level DI tree. */
	private readonly recentColors = inject(RecentColorsService, { optional: true });

	protected readonly applicable = computed(() => hasShapeProperties(this.element()));
	private readonly style = computed<ShapeStyle | undefined>(() => {
		const el = this.element();
		return hasShapeProperties(el) ? el.shapeStyle : undefined;
	});

	/** The shared pattern view; a shape not yet in a pattern fill shows the one it would get. */
	protected readonly pattern = computed<NonNullable<ShapeFillView['pattern']>>(() => {
		const style = this.style();
		const view = resolveShapeFillView(style).pattern;
		return (
			view ?? resolveShapeFillView({ ...style, ...shapeFillModeChange(style, 'pattern') }).pattern!
		);
	});

	protected onPreset(preset: string): void {
		this.act({ type: 'patternPreset', preset });
	}

	protected onColor(event: Event, type: 'patternForeground' | 'patternBackground'): void {
		const target = event.target;
		if (!(target instanceof HTMLInputElement)) {
			return;
		}
		this.act({ type, hex: target.value });
		this.recentColors?.push(target.value);
	}

	private act(action: ShapeFillAction): void {
		const patch = shapeFillPatch(this.style(), action);
		if (patch) {
			this.patch.emit(shapeQuickStylePatch(this.element(), patch));
		}
	}
}
