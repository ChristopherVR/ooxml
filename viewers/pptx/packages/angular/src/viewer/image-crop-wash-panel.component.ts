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
import { cropResetPatch, cropResetState, PICTURE_CROP_SECTIONS } from 'ooxml-ui/pptx';
import type { PptxCropShape, PptxElement, PptxImageEffects } from 'pptx-viewer-core';
import { isImageLikeElement } from 'pptx-viewer-core';

import { InspectorFieldsComponent } from './inspector-fields.component';
import { RecentColorsService } from './recent-colors.service';

const CROP_SHAPES: readonly { value: PptxCropShape; label: string; glyph: string }[] = [
	{ value: 'none', label: 'None', glyph: '▭' },
	{ value: 'ellipse', label: 'Ellipse', glyph: '●' },
	{ value: 'roundedRect', label: 'Rounded rectangle', glyph: '▣' },
	{ value: 'triangle', label: 'Triangle', glyph: '▲' },
	{ value: 'diamond', label: 'Diamond', glyph: '◆' },
	{ value: 'pentagon', label: 'Pentagon', glyph: '⬟' },
	{ value: 'hexagon', label: 'Hexagon', glyph: '⬢' },
	{ value: 'star', label: 'Star', glyph: '★' },
];
type ImageElement = PptxElement & { imageEffects?: PptxImageEffects };

export function replacementImagePatch(dataUrl: string): Partial<PptxElement> {
	return {
		imageData: dataUrl,
		imagePath: undefined,
		svgData: undefined,
		svgPath: undefined,
	} as Partial<PptxElement>;
}

@Component({
	schemas: [CUSTOM_ELEMENTS_SCHEMA],
	selector: 'pptx-image-crop-wash-panel',
	standalone: true,
	imports: [TranslatePipe, InspectorFieldsComponent],
	changeDetection: ChangeDetectionStrategy.OnPush,
	template: `
		<section class="panel" aria-label="Image crop and color wash">
			<label class="replace">
				<span>{{ 'pptx.image.replaceImage' | translate }}</span>
				<input type="file" accept="image/*" (change)="onReplaceImage($event)" />
			</label>
			<pptx-inspector-fields
				[sections]="cropSections"
				[target]="element()"
				[disabled]="!croppable()"
				(patch)="onCrop($event)"
			/>
			<button type="button" [disabled]="!croppable()" (click)="resetCrop()">
				{{ 'pptx.image.resetCrop' | translate }}
			</button>
			<div class="shapes" role="group" aria-label="Crop to shape">
				@for (shape of cropShapes; track shape.value) {
					<button
						type="button"
						[class.active]="cropShape() === shape.value"
						[disabled]="!croppable()"
						[title]="shape.label"
						[attr.aria-label]="shape.label"
						(click)="setCropShape(shape.value)"
					>
						{{ shape.glyph }}
					</button>
				}
			</div>
			<label class="toggle"
				><span>{{ 'pptx.image.colorWash' | translate }}</span
				><pptx-ui-checkbox [checked]="!!wash()" (change)="toggleWash($event)"></pptx-ui-checkbox
			></label>
			@if (wash(); as value) {
				<div class="wash">
					<label
						><span>{{ 'pptx.image.washColor' | translate }}</span
						><input type="color" [value]="value.color || '#0066cc'" (change)="onWashColor($event)"
					/></label>
					<label
						><span>{{ 'pptx.image.washOpacity' | translate }}</span
						><input
							type="range"
							min="0"
							max="100"
							[value]="value.opacity"
							(input)="onWashOpacity($event)"
					/></label>
				</div>
			}
		</section>
	`,
	styles: `
		.panel {
			display: grid;
			gap: 8px;
			padding-top: 4px;
			font-size: 11px;
		}
		.wash {
			display: grid;
			grid-template-columns: repeat(2, minmax(0, 1fr));
			gap: 7px;
		}
		.replace input {
			width: 100%;
			font-size: 10px;
			color: inherit;
		}
		.shapes {
			display: grid;
			grid-template-columns: repeat(4, minmax(0, 1fr));
			gap: 4px;
		}
		.shapes button {
			font-size: 17px;
			line-height: 1;
		}
		.shapes button.active {
			border-color: var(--pptx-primary, #2563eb);
			background: color-mix(in srgb, var(--pptx-primary, #2563eb) 25%, transparent);
		}
		label {
			display: grid;
			gap: 3px;
			color: var(--pptx-inspector-muted, #aaa);
		}
		.toggle {
			display: flex;
			align-items: center;
			justify-content: space-between;
		}
		input[type='range'] {
			width: 100%;
			accent-color: var(--pptx-primary, #2563eb);
		}
		button {
			padding: 4px 7px;
			border: 1px solid var(--pptx-inspector-border, #444);
			border-radius: 3px;
			background: var(--pptx-inspector-input-bg, #2d2d2d);
			color: inherit;
			cursor: pointer;
		}
	`,
})
export class ImageCropWashPanelComponent {
	readonly element = input.required<PptxElement>();
	readonly canEdit = input<boolean>(true);
	readonly patch = output<Partial<PptxElement>>();
	/** Optional: absent in a standalone unit test with no viewer-level DI tree. */
	private readonly recentColors = inject(RecentColorsService, { optional: true });
	protected readonly cropSections = PICTURE_CROP_SECTIONS;
	protected readonly cropShapes = CROP_SHAPES;
	protected readonly effects = computed(() => (this.element() as ImageElement).imageEffects ?? {});
	protected readonly wash = computed(() => this.effects().colorWash);
	protected readonly cropShape = computed(
		() => (this.element() as ImageElement & { cropShape?: PptxCropShape }).cropShape ?? 'none',
	);
	/** G7: `a:picLocks/@noCrop` forbids cropping this specific picture. */
	protected readonly croppable = computed(
		() => cropResetState(this.element(), this.canEdit(), true).enabled,
	);

	protected onReplaceImage(event: Event): void {
		const file = (event.target as HTMLInputElement).files?.[0];
		if (!file) {
			return;
		}
		const reader = new FileReader();
		reader.onload = () => {
			if (typeof reader.result === 'string') {
				this.patch.emit(replacementImagePatch(reader.result));
			}
		};
		reader.readAsDataURL(file);
	}

	protected onCrop(patch: Partial<PptxElement>): void {
		if (isImageLikeElement(this.element()) && this.croppable()) {
			this.patch.emit(patch);
		}
	}
	protected resetCrop(): void {
		if (!this.croppable()) {
			return;
		}
		this.patch.emit(cropResetPatch());
	}
	protected setCropShape(value: PptxCropShape): void {
		if (!this.croppable()) {
			return;
		}
		this.patch.emit({ cropShape: value } as Partial<PptxElement>);
	}
	protected toggleWash(event: Event): void {
		this.updateEffects({
			colorWash: (event.target as HTMLInputElement).checked
				? { color: '#0066cc', opacity: 40 }
				: undefined,
		});
	}
	protected onWashColor(event: Event): void {
		const color = (event.target as HTMLInputElement).value;
		this.updateEffects({
			colorWash: {
				color,
				opacity: this.wash()?.opacity ?? 40,
			},
		});
		this.recentColors?.push(color);
	}
	protected onWashOpacity(event: Event): void {
		this.updateEffects({
			colorWash: {
				color: this.wash()?.color ?? '#0066cc',
				opacity: Number((event.target as HTMLInputElement).value),
			},
		});
	}
	private updateEffects(update: Partial<PptxImageEffects>): void {
		this.patch.emit({ imageEffects: { ...this.effects(), ...update } } as Partial<PptxElement>);
	}
}
