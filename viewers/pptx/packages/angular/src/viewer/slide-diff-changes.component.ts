/**
 * slide-diff-changes.component.ts: the per-element change list for one slide
 * diff.
 *
 * Selector: `pptx-slide-diff-changes`
 *
 * Split out of {@link SlideDiffRowComponent} to keep each file focused. Renders
 * each {@link ElementChange} as an icon-prefixed description row, colouring the
 * icon by change kind.
 */

import { ChangeDetectionStrategy, Component, inject, input } from '@angular/core';
import { TranslateService } from '@ngx-translate/core';
import { describeElementChange, type ElementChange } from 'ooxml-ui/pptx';

import { changeIcon } from './slide-diff-helpers';

@Component({
	selector: 'pptx-slide-diff-changes',
	standalone: true,
	changeDetection: ChangeDetectionStrategy.OnPush,
	template: `
		<div class="pptx-ng-diff-changes">
			@for (change of changes(); track $index) {
				<div class="pptx-ng-diff-change">
					<span class="pptx-ng-diff-change-icon" [attr.data-kind]="change.kind">
						{{ icon(change.kind) }}
					</span>
					<span class="pptx-ng-diff-change-desc">{{ describe(change) }}</span>
				</div>
			}
		</div>
	`,
	styles: [
		`
			.pptx-ng-diff-changes {
				display: flex;
				flex-direction: column;
				gap: 0.25rem;
			}
			.pptx-ng-diff-change {
				display: flex;
				align-items: flex-start;
				gap: 0.5rem;
				padding: 0.375rem 0.5rem;
				border-radius: 0.25rem;
				background: rgba(107, 114, 128, 0.15);
				font-size: 0.6875rem;
			}
			.pptx-ng-diff-change-icon {
				flex-shrink: 0;
				font-weight: 700;
				color: var(--pptx-muted-foreground, #9ca3af);
			}
			.pptx-ng-diff-change-icon[data-kind='added'] {
				color: #4ade80;
			}
			.pptx-ng-diff-change-icon[data-kind='removed'] {
				color: #f87171;
			}
			.pptx-ng-diff-change-icon[data-kind='moved'] {
				color: var(--pptx-primary, #6366f1);
			}
			.pptx-ng-diff-change-icon[data-kind='resized'] {
				color: #fbbf24;
			}
			.pptx-ng-diff-change-icon[data-kind='textChanged'] {
				color: #c084fc;
			}
			.pptx-ng-diff-change-desc {
				color: var(--pptx-foreground, #f3f4f6);
			}
		`,
	],
})
export class SlideDiffChangesComponent {
	readonly changes = input.required<readonly ElementChange[]>();

	private readonly translate = inject(TranslateService);
	protected readonly icon = changeIcon;

	protected describe(change: ElementChange): string {
		return describeElementChange(change, (key, params) => this.translate.instant(key, params));
	}
}
