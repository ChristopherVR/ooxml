/**
 * StrokeOutlineComponent: the stroked SVG outline from shared
 * `buildStrokeOutline` (a gradient or pattern `a:ln`, a stroke-only preset, or
 * a centred flat line), projected into the host `<svg>`.
 *
 * Selector: `svg[pptx-stroke-outline]` (attribute selector, so the host stays
 * the plain `<svg>` the caller positions and gives its `viewBox`). Shapes and
 * pictures share it, so a picture's frame paints exactly like a shape's line.
 */
import { ChangeDetectionStrategy, Component, Input } from '@angular/core';
import type { StrokeOutline } from 'ooxml-ui/pptx';

@Component({
	selector: 'svg[pptx-stroke-outline]',
	standalone: true,
	changeDetection: ChangeDetectionStrategy.OnPush,
	template: `
		@let o = outline;
		@if (o.paint; as paint) {
			<svg:defs>
				@if (paint.kind === 'pattern') {
					<svg:pattern
						[attr.id]="paint.id"
						[attr.width]="paint.width"
						[attr.height]="paint.height"
						patternUnits="userSpaceOnUse"
					>
						<svg:image
							[attr.href]="paint.href"
							[attr.width]="paint.width"
							[attr.height]="paint.height"
						/>
					</svg:pattern>
				} @else if (paint.kind === 'rectPath') {
					<svg:pattern [attr.id]="paint.id" patternUnits="objectBoundingBox" width="1" height="1">
						<svg:image
							[attr.href]="paint.href"
							x="0"
							y="0"
							width="1"
							height="1"
							preserveAspectRatio="none"
						/>
					</svg:pattern>
				} @else if (paint.kind === 'radial') {
					<svg:radialGradient
						[attr.id]="paint.id"
						[attr.cx]="paint.cx"
						[attr.cy]="paint.cy"
						[attr.r]="paint.r"
					>
						@for (stop of paint.stops; track $index) {
							<svg:stop
								[attr.offset]="stop.offset"
								[attr.stop-color]="stop.color"
								[attr.stop-opacity]="stop.opacity"
							/>
						}
					</svg:radialGradient>
				} @else {
					<svg:linearGradient
						[attr.id]="paint.id"
						[attr.x1]="paint.x1"
						[attr.y1]="paint.y1"
						[attr.x2]="paint.x2"
						[attr.y2]="paint.y2"
					>
						@for (stop of paint.stops; track $index) {
							<svg:stop
								[attr.offset]="stop.offset"
								[attr.stop-color]="stop.color"
								[attr.stop-opacity]="stop.opacity"
							/>
						}
					</svg:linearGradient>
				}
			</svg:defs>
		}
		@for (strand of o.strands; track $index) {
			<svg:path
				[attr.d]="o.d"
				fill="none"
				[attr.stroke]="o.stroke"
				[attr.stroke-width]="strand.strokeWidth"
				[style.stroke-width]="strand.cssStrokeWidth"
				[attr.stroke-dasharray]="o.dashArray"
				[attr.stroke-linecap]="o.lineCap"
				[attr.stroke-linejoin]="o.lineJoin"
				[style.transform]="strand.offset !== 0 ? 'translate(0, ' + strand.offset + 'px)' : null"
			/>
		}
	`,
})
export class StrokeOutlineComponent {
	// A decorator input, not a signal `input()`: the host is matched by an
	// attribute selector inside other components' templates, and plain JIT
	// (the viewer's TestBed specs) only binds decorator inputs.
	@Input({ required: true }) outline!: StrokeOutline;
}
