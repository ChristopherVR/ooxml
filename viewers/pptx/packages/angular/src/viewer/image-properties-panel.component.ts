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
	ARTISTIC_EFFECTS,
	DUOTONE_PRESETS,
	ELEMENT_ALT_TEXT_SECTIONS,
	imageResetPatch,
	imageResetState,
	PICTURE_ADJUSTMENT_SECTIONS,
	PICTURE_EFFECT_SECTIONS,
	pictureDuotonePresetPatch,
} from 'ooxml-ui/pptx';
import type { DuotonePreset } from 'ooxml-ui/pptx';
import type { PptxElement, PptxImageEffects } from 'pptx-viewer-core';
import { isImageLikeElement } from 'pptx-viewer-core';

import { ImageCropWashPanelComponent } from './image-crop-wash-panel.component';
import { InspectorFieldsComponent } from './inspector-fields.component';

type ImageElement = PptxElement & { imageEffects?: PptxImageEffects };

export function mergeImageEffects(
	effects: PptxImageEffects | undefined,
	update: Partial<PptxImageEffects>,
): PptxImageEffects {
	return { ...(effects ?? {}), ...update };
}

@Component({
	schemas: [CUSTOM_ELEMENTS_SCHEMA],
	selector: 'pptx-image-properties-panel',
	standalone: true,
	imports: [TranslatePipe, ImageCropWashPanelComponent, InspectorFieldsComponent],
	changeDetection: ChangeDetectionStrategy.OnPush,
	template: `
		<div class="pptx-ng-image-properties">
			<pptx-inspector-fields
				[sections]="altTextSections"
				[target]="element()"
				[disabled]="!canEdit()"
				(patch)="patch.emit($event)"
			/>
			<pptx-inspector-fields
				[sections]="adjustmentSections"
				[target]="element()"
				[disabled]="!canEdit()"
				(patch)="patch.emit($event)"
			/>
			<label class="field field--stacked">
				<span>{{ 'pptx.image.artisticEffects' | translate }}</span>
				<pptx-ui-select
					[attr.aria-label]="'pptx.image.artisticEffects' | translate"
					[value]="effects().artisticEffect ?? 'none'"
					(change)="onArtistic($event)"
				>
					@for (effect of artisticEffects; track effect[0]) {
						<option
							[value]="effect[0]"
							[selected]="effect[0] === (effects().artisticEffect ?? 'none')"
						>
							{{ effect[1] | translate }}
						</option>
					}
				</pptx-ui-select>
			</label>
			<pptx-inspector-fields
				[sections]="effectSections"
				[target]="element()"
				[disabled]="!canEdit()"
				(patch)="patch.emit($event)"
			/>
			<div class="duotone-presets">
				<span>{{ 'pptx.image.duotonePresets' | translate }}</span>
				@for (preset of duotonePresets; track preset.labelKey) {
					<button
						type="button"
						class="duotone-preset"
						data-testid="picture-duotone-preset"
						[disabled]="!canEdit()"
						[title]="preset.labelKey | translate"
						[attr.aria-label]="preset.labelKey | translate"
						[style.background]="
							'linear-gradient(135deg, ' + preset.shadow + ' 50%, ' + preset.highlight + ' 50%)'
						"
						(click)="patch.emit(duotonePreset(preset))"
					></button>
				}
			</div>
			<pptx-image-crop-wash-panel
				[element]="element()"
				[canEdit]="canEdit()"
				(patch)="patch.emit($event)"
			/>
			<button type="button" class="reset" [disabled]="!resetEnabled()" (click)="reset()">
				{{ 'pptx.image.resetImage' | translate }}
			</button>
		</div>
	`,
	styles: `
		.pptx-ng-image-properties {
			display: grid;
			gap: 9px;
			font-size: 11px;
		}
		.field {
			display: grid;
			grid-template-columns: 78px 1fr 30px;
			align-items: center;
			gap: 6px;
			color: var(--pptx-inspector-muted, #aaa);
		}
		.field--stacked {
			grid-template-columns: 1fr;
		}
		input[type='text'],
		select {
			box-sizing: border-box;
			width: 100%;
			padding: 4px 6px;
			border: 1px solid var(--pptx-inspector-border, #444);
			border-radius: 3px;
			background: var(--pptx-inspector-input-bg, #2d2d2d);
			color: inherit;
		}
		input[type='range'] {
			min-width: 0;
			accent-color: var(--pptx-primary, #2563eb);
		}
		.duotone-presets {
			display: flex;
			flex-wrap: wrap;
			align-items: center;
			gap: 4px;
			color: var(--pptx-inspector-muted, #aaa);
		}
		.duotone-presets > span {
			width: 100%;
		}
		.duotone-preset {
			width: 24px;
			height: 24px;
			padding: 0;
			border-radius: 50%;
		}
		button {
			padding: 4px 7px;
			border: 1px solid var(--pptx-inspector-border, #444);
			border-radius: 3px;
			background: var(--pptx-inspector-input-bg, #2d2d2d);
			color: inherit;
			cursor: pointer;
		}
		.reset {
			width: 100%;
		}
	`,
})
export class ImagePropertiesPanelComponent {
	readonly element = input.required<PptxElement>();
	readonly canEdit = input<boolean>(true);
	readonly patch = output<Partial<PptxElement>>();
	protected readonly resetEnabled = computed(
		() => imageResetState(this.element(), this.canEdit(), true).enabled,
	);

	protected readonly image = computed(() => this.element() as ImageElement);
	protected readonly effects = computed(() => this.image().imageEffects ?? {});
	protected readonly artisticEffects = ARTISTIC_EFFECTS;
	protected readonly altTextSections = ELEMENT_ALT_TEXT_SECTIONS;
	protected readonly adjustmentSections = PICTURE_ADJUSTMENT_SECTIONS;
	protected readonly effectSections = PICTURE_EFFECT_SECTIONS;
	protected readonly duotonePresets = DUOTONE_PRESETS;

	static supports(element: PptxElement): boolean {
		return isImageLikeElement(element);
	}

	private updateEffects(update: Partial<PptxImageEffects>): void {
		this.patch.emit({
			imageEffects: mergeImageEffects(this.effects(), update),
		} as Partial<PptxElement>);
	}

	protected onArtistic(event: Event): void {
		const value = (event.target as HTMLSelectElement).value;
		this.updateEffects({ artisticEffect: value === 'none' ? undefined : value });
	}

	protected duotonePreset(preset: DuotonePreset): Partial<PptxElement> {
		return pictureDuotonePresetPatch(this.element(), preset);
	}

	protected reset(): void {
		if (this.resetEnabled()) {
			this.patch.emit(imageResetPatch());
		}
	}
}
