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
import type { RibbonGalleryPickEvent } from 'ooxml-ui/pptx';
import { buildTextWarpGallery, TEXT_WARP_PRESETS, textWarpStyle } from 'ooxml-ui/pptx';
import type { PptxElement, PptxTextWarpPreset, TextStyle } from 'pptx-viewer-core';
import { hasTextProperties } from 'pptx-viewer-core';

import { homeLanguage, ngxTranslate } from './ribbon-home-lang';

export function textWarpPatch(
	element: PptxElement,
	value: PptxTextWarpPreset,
): Partial<PptxElement> {
	const current = (element as PptxElement & { textStyle?: TextStyle }).textStyle;
	return { textStyle: { ...current, ...textWarpStyle(value) } } as Partial<PptxElement>;
}

@Component({
	selector: 'pptx-text-warp-gallery',
	standalone: true,
	imports: [TranslatePipe],
	changeDetection: ChangeDetectionStrategy.OnPush,
	schemas: [CUSTOM_ELEMENTS_SCHEMA],
	template: `
		@if (supported()) {
			<details class="panel">
				<summary>
					{{ 'pptx.textWarp.title' | translate }} <span>{{ currentLabel() | translate }}</span>
				</summary>
				<pptx-ui-ribbon-gallery
					class="gallery"
					mode="panel"
					compact
					[descriptor]="descriptor()"
					[translateLabel]="translator()"
					(gallery-pick)="onPick($event)"
				/>
			</details>
		}
	`,
	styles: `
		.panel {
			padding: 8px 0;
			border-bottom: 1px solid var(--pptx-inspector-border, #333);
			font-size: 11px;
		}
		summary {
			display: flex;
			justify-content: space-between;
			gap: 8px;
			cursor: pointer;
			color: var(--pptx-inspector-muted, #aaa);
		}
		summary span {
			overflow: hidden;
			text-overflow: ellipsis;
			white-space: nowrap;
			text-transform: none;
		}
		.gallery {
			margin-top: 8px;
		}
	`,
})
export class TextWarpGalleryComponent {
	private readonly translation = inject(TranslateService, { optional: true });
	/** Read so a language switch re-labels the tiles. */
	private readonly language = homeLanguage(this.translation);

	readonly element = input.required<PptxElement>();
	readonly patch = output<Partial<PptxElement>>();
	protected readonly supported = computed(() => hasTextProperties(this.element()));
	protected readonly current = computed<PptxTextWarpPreset>(() =>
		hasTextProperties(this.element())
			? (this.element() as PptxElement & { textStyle?: TextStyle }).textStyle?.textWarpPreset ||
				'textNoShape'
			: 'textNoShape',
	);
	protected readonly currentLabel = computed(
		() =>
			TEXT_WARP_PRESETS.find((preset) => preset.value === this.current())?.labelKey ??
			this.current(),
	);
	protected readonly descriptor = computed(() => buildTextWarpGallery(this.current()));
	protected readonly translator = computed(() => {
		this.language();
		return ngxTranslate(this.translation);
	});

	protected onPick(event: Event): void {
		const { itemId } = (event as RibbonGalleryPickEvent).detail;
		if (textWarpStyle(itemId)) {
			this.patch.emit(textWarpPatch(this.element(), itemId as PptxTextWarpPreset));
		}
	}
}
