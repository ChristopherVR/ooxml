/**
 * ribbon-picture-size.component.ts: Picture Format > Size, a thin adapter
 * around `pptx-ui-ribbon-home-picture-size`. Its Crop split button runs the
 * same crop commands as Home > Arrange (`ribbon-crop-commands.ts`).
 */
import {
	ChangeDetectionStrategy,
	Component,
	CUSTOM_ELEMENTS_SCHEMA,
	inject,
	input,
} from '@angular/core';
import { TranslateService } from '@ngx-translate/core';
import { pictureSizeHomeControls } from 'ooxml-ui/pptx';
import type { RibbonHomeRequestEvent, ToolbarActionId } from 'ooxml-ui/pptx';
import type { PptxElement } from 'pptx-viewer-core';

import { EditorStateService } from './editor-state.service';
import { LoadContentService } from './load-content.service';
import { PictureCropService } from './picture-crop.service';
import { createRibbonCropCommands } from './ribbon-crop-commands';
import { homeLanguage, homeTranslator } from './ribbon-home-lang';

@Component({
	selector: 'pptx-ribbon-picture-size',
	standalone: true,
	changeDetection: ChangeDetectionStrategy.OnPush,
	host: { class: 'contents' },
	schemas: [CUSTOM_ELEMENTS_SCHEMA],
	template: `<pptx-ui-ribbon-home-picture-size
		[state]="view()"
		(home-request)="request($event)"
	/>`,
})
export class RibbonPictureSizeComponent {
	private readonly translation = inject(TranslateService, { optional: true });
	private readonly language = homeLanguage(this.translation);

	readonly slideIndex = input<number>(0);
	readonly selectedElement = input<PptxElement | null>(null);
	readonly canEdit = input<boolean>(false);
	/** Toolbar buttons the host hides (gates Crop). */
	readonly hiddenActions = input<readonly ToolbarActionId[]>([]);

	private readonly cropCommands = createRibbonCropCommands({
		editor: inject(EditorStateService),
		crop: inject(PictureCropService, { optional: true }),
		loader: inject(LoadContentService, { optional: true }),
		slideIndex: () => this.slideIndex(),
		selectedElement: () => this.selectedElement(),
		canEdit: () => this.canEdit(),
	});

	protected view() {
		return {
			controls: pictureSizeHomeControls({
				editable: this.canEdit(),
				canCrop: this.cropCommands.target() !== null,
				cropActive: this.cropCommands.active(),
				hideCrop: this.hiddenActions().includes('crop'),
			}),
			translate: homeTranslator(this.translation, this.language, ['picture-size']),
		};
	}

	protected request(event: Event): void {
		this.cropCommands.run((event as RibbonHomeRequestEvent).detail.value);
	}
}
