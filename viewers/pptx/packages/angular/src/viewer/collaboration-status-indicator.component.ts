/**
 * collaboration-status-indicator.component.ts: the status bar's collaboration
 * pill (colour dot, state text, Retry in the error state). Angular port of the
 * Vue `CollaborationStatusIndicator.vue`; the viewer projects it into the
 * shared status bar's `collaboration` slot. In the error state its tooltip is
 * the sentence `collaborationFailureMessageKey` names for the failure.
 */
import { ChangeDetectionStrategy, Component, computed, input, output } from '@angular/core';
import { TranslatePipe } from '@ngx-translate/core';
import { collaborationFailureMessageKey } from 'ooxml-ui/pptx';
import type { CollaborationFailure, ConnectionStatus } from 'ooxml-ui/pptx';

interface StatusStyle {
	dot: string;
	text: string;
	labelKey: string;
}

const STATUS_STYLES: Record<ConnectionStatus, StatusStyle> = {
	connected: {
		dot: 'bg-green-400',
		text: 'text-green-400',
		labelKey: 'pptx.collaboration.status.connected',
	},
	connecting: {
		dot: 'bg-yellow-400 animate-pulse',
		text: 'text-yellow-400',
		labelKey: 'pptx.collaboration.status.connecting',
	},
	disconnected: {
		dot: 'bg-gray-500',
		text: 'text-gray-500',
		labelKey: 'pptx.collaboration.status.disconnected',
	},
	error: { dot: 'bg-red-400', text: 'text-red-400', labelKey: 'pptx.collaboration.status.error' },
};

@Component({
	selector: 'pptx-collaboration-status',
	standalone: true,
	changeDetection: ChangeDetectionStrategy.OnPush,
	imports: [TranslatePipe],
	template: `
		<div
			role="status"
			data-testid="collaboration-status"
			class="flex items-center gap-1.5"
			[attr.aria-label]="
				'pptx.collaboration.statusAriaLabel'
					| translate: { status: style().labelKey | translate, count: connectedCount() }
			"
			[attr.title]="failureKey() ? (failureKey()! | translate) : null"
		>
			<span
				class="inline-block h-2 w-2 rounded-full"
				[class]="style().dot"
				aria-hidden="true"
			></span>
			<span class="text-[10px]" [class]="style().text">
				@if (status() === 'connected') {
					@if (connectedCount() === 1) {
						{{ 'pptx.collaboration.onePersonHere' | translate }}
					} @else {
						{{ 'pptx.collaboration.peopleHere' | translate: { count: connectedCount() } }}
					}
				} @else {
					{{ style().labelKey | translate }}
				}
			</span>
			@if (status() === 'error') {
				<button
					type="button"
					class="text-[10px] text-blue-400 underline underline-offset-2 transition-colors hover:text-blue-300"
					[attr.aria-label]="'pptx.collaboration.retry' | translate"
					(click)="retry.emit()"
				>
					{{ 'pptx.collaboration.retry' | translate }}
				</button>
			}
		</div>
	`,
})
export class CollaborationStatusIndicatorComponent {
	/** Current connection status. */
	readonly status = input<ConnectionStatus>('disconnected');
	/** Connected participants, the local user included. */
	readonly connectedCount = input<number>(0);
	/** Why the session failed to start; its sentence is the tooltip in the error state. */
	readonly failure = input<CollaborationFailure | null>(null);
	/** The user asked to retry after a connection error. */
	readonly retry = output<void>();

	protected readonly style = computed(() => STATUS_STYLES[this.status()]);
	protected readonly failureKey = computed(() =>
		collaborationFailureMessageKey(this.status(), this.failure()),
	);
}
