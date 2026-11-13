import { ChangeDetectionStrategy, Component, computed, input } from '@angular/core';
import { describeBuildingBlockToolbar } from 'ooxml-ui/pptx';
import type { BuildingBlockToolbarGroupId } from 'ooxml-ui/pptx';

import type { ViewerBuildingBlocks } from './viewer-building-blocks';

/** The compact toolbar block: one button per control of `describeBuildingBlockToolbar`. */
@Component({
	selector: 'pptx-building-blocks-toolbar',
	standalone: true,
	changeDetection: ChangeDetectionStrategy.OnPush,
	host: { 'data-pptx-block': 'toolbar', role: 'toolbar' },
	styles: `
		:host {
			display: flex;
			flex-wrap: wrap;
			gap: 4px 12px;
		}
	`,
	template: `
		@for (group of groups(); track group.id) {
			<span role="group" [attr.data-group]="group.id">
				@for (control of group.controls; track control.id) {
					<button
						type="button"
						[attr.data-control]="control.id"
						[disabled]="control.disabled"
						[attr.aria-pressed]="control.pressed === undefined ? null : control.pressed"
						(click)="blocks().runToolbar(control.id)"
					>
						{{ blocks().translate(control.labelKey) }}
					</button>
				}
			</span>
		}
	`,
})
export class BuildingBlocksToolbarComponent {
	readonly blocks = input.required<ViewerBuildingBlocks>();
	/** Show only these groups (default: all). */
	readonly include = input<readonly BuildingBlockToolbarGroupId[] | undefined>(undefined);

	protected readonly groups = computed(() => {
		const blocks = this.blocks();
		const state = blocks.state();
		const groups = describeBuildingBlockToolbar({
			availability: blocks.actions().availability,
			canUndo: state.canUndo,
			canRedo: state.canRedo,
			findOpen: blocks.findBarState().open,
			slideIndex: state.slideIndex,
			slideCount: state.slideCount,
			hiddenActions: blocks.options().hiddenActions,
		});
		const include = this.include();
		return include ? groups.filter((group) => include.includes(group.id)) : groups;
	});
}
