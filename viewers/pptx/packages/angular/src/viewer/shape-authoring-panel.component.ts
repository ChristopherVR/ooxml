import {
	ChangeDetectionStrategy,
	Component,
	computed,
	input,
	output,
	CUSTOM_ELEMENTS_SCHEMA,
	inject,
} from '@angular/core';
import { TranslatePipe, TranslateService } from '@ngx-translate/core';
import {
	buildQuickStylesGallery,
	clampShapeAdjustmentValue,
	DEFAULT_ROUND_RECT_ADJUSTMENT,
	quickStyleShapeStyle,
	SHAPE_ADJUSTMENT_MAX,
	SHAPE_ADJUSTMENT_MIN,
	SHAPE_PRESET_DEFS,
	shapeQuickStylePatch,
} from 'ooxml-ui/pptx';
import type { RibbonGalleryPickEvent } from 'ooxml-ui/pptx';
import type { PptxElement } from 'pptx-viewer-core';
import { hasShapeProperties } from 'pptx-viewer-core';

import { homeLanguage, ngxTranslate } from './ribbon-home-lang';

export function shapeTypePatch(element: PptxElement, shapeType: string): Partial<PptxElement> {
	if (!hasShapeProperties(element)) {
		return {};
	}
	const existing = { ...element.shapeAdjustments };
	const adjustment =
		shapeType === 'roundRect'
			? { ...existing, adj: existing.adj ?? DEFAULT_ROUND_RECT_ADJUSTMENT }
			: shapeType === 'cylinder' || shapeType === 'can'
				? { ...existing, adj: existing.adj ?? 25000 }
				: undefined;
	const line = shapeType === 'line';
	return {
		shapeType,
		shapeAdjustments: adjustment,
		shapeStyle: {
			...element.shapeStyle,
			fillColor: line ? 'transparent' : element.shapeStyle?.fillColor || '#4472c4',
			fillMode: line ? 'none' : element.shapeStyle?.fillMode || 'solid',
			strokeWidth: line
				? Math.max(2, element.shapeStyle?.strokeWidth || 0)
				: element.shapeStyle?.strokeWidth || 1,
		},
	} as Partial<PptxElement>;
}

@Component({
	schemas: [CUSTOM_ELEMENTS_SCHEMA],
	selector: 'pptx-shape-authoring-panel',
	standalone: true,
	imports: [TranslatePipe],
	changeDetection: ChangeDetectionStrategy.OnPush,
	template: `
		@if (supported()) {
			<section class="panel" aria-label="Shape authoring">
				<label class="field">
					<span>Shape type</span>
					<pptx-ui-select
						aria-label="Shape type"
						[value]="shapeType()"
						(change)="onShapeType($event)"
					>
						@for (preset of presets; track preset.type) {
							<option [value]="preset.type" [selected]="preset.type === shapeType()">
								{{ preset.label }}
							</option>
						}
						<!--
							Accepted deviation: this prints the raw a:prstGeom token. The deck
							may carry any of the 188 presets while the gallery offers a subset,
							and showing the token is what keeps the current geometry
							representable, so the picker cannot silently rewrite it. There is
							no catalogue of preset-geometry names to spell it from.
						-->
						@if (unknownType()) {
							<option [value]="shapeType()" [selected]="true">{{ shapeType() }}</option>
						}
					</pptx-ui-select>
				</label>
				@if (showAdjustment()) {
					<label class="field">
						<span>Roundness</span>
						<input
							type="range"
							[min]="adjustmentMin"
							[max]="adjustmentMax"
							step="500"
							[value]="adjustment()"
							(input)="onAdjustment($event)"
						/>
					</label>
				}
				<!-- A div, not a label: a label would hand a click on its caption to the first tile. -->
				<div class="field">
					<span>{{ 'pptx.shape.quickStyles' | translate }}</span>
					<pptx-ui-ribbon-gallery
						mode="panel"
						compact
						[descriptor]="quickStyles"
						[translateLabel]="translator()"
						(gallery-pick)="onQuickStyle($event)"
					/>
				</div>
			</section>
		}
	`,
	styles: `
		.panel {
			display: grid;
			gap: 8px;
			padding: 8px 0;
			border-bottom: 1px solid var(--pptx-inspector-border, #333);
			font-size: 11px;
		}
		.field {
			display: grid;
			gap: 4px;
			color: var(--pptx-inspector-muted, #aaa);
		}
		select {
			width: 100%;
			padding: 4px;
			border: 1px solid var(--pptx-inspector-border, #444);
			border-radius: 3px;
			background: var(--pptx-inspector-input-bg, #2d2d2d);
			color: inherit;
		}
		input {
			width: 100%;
			accent-color: var(--pptx-primary, #2563eb);
		}
		/* The Shape type select gets the WCAG touch target below the mobile
		   breakpoint (768px); the Quick Styles gallery sizes its own tiles. */
		@media (max-width: 767px) {
			select {
				min-height: 44px;
			}
		}
	`,
})
export class ShapeAuthoringPanelComponent {
	readonly element = input.required<PptxElement>();
	readonly patch = output<Partial<PptxElement>>();
	protected readonly presets = SHAPE_PRESET_DEFS.filter((preset) => preset.type !== 'connector');
	/** The shared Quick Styles gallery: one tile per shape style preset. */
	protected readonly quickStyles = buildQuickStylesGallery();
	private readonly translation = inject(TranslateService, { optional: true });
	/** Read so a language switch re-labels the presets. */
	private readonly language = homeLanguage(this.translation);
	protected readonly translator = computed(() => {
		this.language();
		return ngxTranslate(this.translation);
	});
	protected readonly adjustmentMin = SHAPE_ADJUSTMENT_MIN;
	protected readonly adjustmentMax = SHAPE_ADJUSTMENT_MAX;
	protected readonly supported = computed(() => ['shape', 'text'].includes(this.element().type));
	protected readonly shapeType = computed(() =>
		hasShapeProperties(this.element())
			? (this.element() as PptxElement & { shapeType?: string }).shapeType || 'rect'
			: 'rect',
	);
	protected readonly unknownType = computed(
		() => !this.presets.some((preset) => preset.type === this.shapeType()),
	);
	protected readonly showAdjustment = computed(() =>
		['roundRect', 'cylinder', 'can'].includes(this.shapeType()),
	);
	protected readonly adjustment = computed(() =>
		hasShapeProperties(this.element())
			? ((this.element() as PptxElement & { shapeAdjustments?: Record<string, number> })
					.shapeAdjustments?.adj ?? DEFAULT_ROUND_RECT_ADJUSTMENT)
			: DEFAULT_ROUND_RECT_ADJUSTMENT,
	);

	protected onShapeType(event: Event): void {
		this.patch.emit(shapeTypePatch(this.element(), (event.target as HTMLSelectElement).value));
	}
	protected onAdjustment(event: Event): void {
		const value = clampShapeAdjustmentValue(Number((event.target as HTMLInputElement).value));
		this.patch.emit({
			shapeAdjustments: {
				...(hasShapeProperties(this.element())
					? (this.element() as PptxElement & { shapeAdjustments?: Record<string, number> })
							.shapeAdjustments
					: {}),
				adj: value,
			},
		} as Partial<PptxElement>);
	}
	/** A Quick Styles pick merges the preset over the shape's own style. */
	protected onQuickStyle(event: Event): void {
		const style = quickStyleShapeStyle((event as RibbonGalleryPickEvent).detail.itemId);
		if (style) {
			this.patch.emit(shapeQuickStylePatch(this.element(), style));
		}
	}
}
