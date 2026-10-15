/**
 * mobile-insert-sheet.component.ts: the bottom bar's Insert sheet. The shared
 * `pptx-ui-mobile-insert` tile grid inside a mobile bottom sheet, routed through the same
 * {@link RibbonInsertSectionComponent} handling as the ribbon, so a pick runs the identical
 * command (React: `MobileInsertSheet.tsx`).
 */
import { ChangeDetectionStrategy, Component, input, output } from '@angular/core';
import { TranslatePipe } from '@ngx-translate/core';

import { MobileSheetComponent } from './mobile-sheet.component';
import { RibbonInsertSectionComponent } from './ribbon-insert-section.component';

@Component({
	selector: 'pptx-mobile-insert-sheet',
	standalone: true,
	changeDetection: ChangeDetectionStrategy.OnPush,
	imports: [MobileSheetComponent, RibbonInsertSectionComponent, TranslatePipe],
	template: `
		<pptx-mobile-sheet
			[open]="open()"
			[title]="'pptx.mobileBar.insert' | translate"
			(closed)="closed.emit()"
		>
			<div class="p-3">
				<pptx-ribbon-insert-section
					[phone]="true"
					[canEdit]="canEdit()"
					[slideIndex]="slideIndex()"
					[commentsOpen]="commentsOpen()"
					(handled)="closed.emit()"
					(openSmartArtDialog)="openSmartArtDialog.emit()"
					(openEquationDialog)="openEquationDialog.emit()"
					(openHyperlink)="openHyperlink.emit()"
					(toggleComments)="toggleComments.emit()"
				/>
			</div>
		</pptx-mobile-sheet>
	`,
})
export class MobileInsertSheetComponent {
	readonly open = input<boolean>(false);
	readonly canEdit = input<boolean>(false);
	readonly slideIndex = input<number>(0);
	readonly commentsOpen = input<boolean>(false);

	readonly closed = output<void>();
	readonly openSmartArtDialog = output<void>();
	readonly openEquationDialog = output<void>();
	readonly openHyperlink = output<void>();
	readonly toggleComments = output<void>();
}
