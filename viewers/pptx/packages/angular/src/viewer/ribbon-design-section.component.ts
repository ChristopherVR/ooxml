import {
	ChangeDetectionStrategy,
	Component,
	CUSTOM_ELEMENTS_SCHEMA,
	input,
	output,
} from '@angular/core';
import { TranslatePipe } from '@ngx-translate/core';
import {
	DESIGN_RIBBON_COMMANDS,
	DESIGN_RIBBON_GROUPS,
	FIXED_TAB_GALLERIES,
	designCommandState,
} from 'ooxml-ui/pptx';
import type { RibbonControlId } from 'ooxml-ui/pptx';
import type { PptxElement } from 'pptx-viewer-core';

import { RibbonGalleryComponent } from './ribbon-gallery.component';

@Component({
	selector: 'pptx-ribbon-design-section',
	standalone: true,
	changeDetection: ChangeDetectionStrategy.OnPush,
	host: { class: 'contents' },
	schemas: [CUSTOM_ELEMENTS_SCHEMA],
	imports: [TranslatePipe, RibbonGalleryComponent],
	template: `
		@for (group of groups; track group.id) {
			<pptx-ui-ribbon-group
				[attr.label]="group.labelKey | translate"
				[attr.data-ribbon-group]="group.id"
			>
				@for (placement of galleriesFor(group.id); track placement.control) {
					<pptx-ribbon-gallery
						[gallery]="placement.gallery"
						[control]="placement.control"
						[mode]="placement.mode"
						[element]="selectedElement()"
						[slideIndex]="slideIndex()"
						[canEdit]="canEdit()"
					/>
				}
				@for (command of commands; track command.id) {
					@if (command.id.startsWith(group.id + '.')) {
						<pptx-ui-ribbon-command
							[attr.data-ribbon-control]="command.id"
							[attr.label]="command.labelKey | translate"
							[attr.title]="command.titleKey | translate"
							[attr.icon]="command.icon"
							[attr.disabled]="view(command.id).disabled ? '' : null"
							[attr.active]="view(command.id).active ? '' : null"
							[attr.expanded]="view(command.id).expanded"
							(command-request)="request(command.id)"
						/>
					}
				}
			</pptx-ui-ribbon-group>
		}
	`,
})
export class RibbonDesignSectionComponent {
	readonly groups = DESIGN_RIBBON_GROUPS;
	readonly commands = DESIGN_RIBBON_COMMANDS;
	readonly themeGalleryOpen = input(false);
	readonly selectedElement = input<PptxElement | null>(null);
	readonly slideIndex = input(0);
	readonly canEdit = input(false);
	readonly toggleThemeGallery = output<void>();
	readonly editTheme = output<void>();
	readonly openSlideSize = output<void>();
	readonly toggleInspector = output<void>();
	protected galleriesFor(groupId: string) {
		return FIXED_TAB_GALLERIES.filter((placement) => placement.control.startsWith(`${groupId}.`));
	}
	protected view(id: RibbonControlId) {
		return designCommandState(id, {
			editable: this.canEdit(),
			galleryOpen: this.themeGalleryOpen(),
		});
	}
	protected request(id: RibbonControlId): void {
		if (this.view(id).disabled) {
			return;
		}
		if (id === 'design.themes.browseThemes') {
			this.toggleThemeGallery.emit();
		} else if (id === 'design.themes.editTheme') {
			this.editTheme.emit();
		} else if (id === 'design.customize.slideSize') {
			this.openSlideSize.emit();
		} else if (id === 'design.customize.formatBackground') {
			this.toggleInspector.emit();
		}
	}
}
