/**
 * signature-status-badge.component.ts: the status bar entry of a deck's digital
 * signatures. It shows only when the deck carries signature parts, and reads as
 * signed only when they verify: label, tooltip, tone and icon all come from the
 * shared `signatureStatusView` in the viewer's current language. Clicking it
 * opens the signatures panel.
 *
 * Selector: `pptx-signature-status-badge`
 */
import { ChangeDetectionStrategy, Component, computed, inject, input, output } from '@angular/core';
import { LucideDynamicIcon } from '@lucide/angular';
import { TranslateService } from '@ngx-translate/core';
import { signatureStatusView } from 'ooxml-ui/pptx';
import type { SignatureStatusView, SignatureTone } from 'ooxml-ui/pptx';
import type { SignatureSummary } from 'pptx-viewer-core';

import { SIGNATURE_ICONS } from './signature-icons';

const TONE_CLASS: Readonly<Record<SignatureTone, string>> = {
	success: 'text-green-400',
	caution: 'text-amber-400',
	danger: 'text-red-400',
	neutral: 'text-gray-400',
};

@Component({
	selector: 'pptx-signature-status-badge',
	standalone: true,
	changeDetection: ChangeDetectionStrategy.OnPush,
	imports: [LucideDynamicIcon],
	template: `
		@if (view().present) {
			<button
				type="button"
				class="pptx-ng-signature-badge inline-flex items-center gap-1 rounded px-1.5 text-[10px] hover:bg-white/10"
				[class]="tone()"
				data-testid="signature-status-badge"
				[attr.data-pptx-signature-state]="view().state"
				[attr.title]="view().description"
				[attr.aria-label]="view().description"
				(click)="open.emit()"
			>
				<svg class="h-3 w-3 flex-none" aria-hidden="true" [lucideIcon]="icons[view().icon]"></svg>
				<span>{{ view().label }}</span>
			</button>
		}
	`,
})
export class SignatureStatusBadgeComponent {
	private readonly translate = inject(TranslateService);

	/** The loaded deck's signatures (`PptxData.signatures`); absent for an unsigned deck. */
	readonly summary = input<SignatureSummary | undefined>(undefined);
	/** The user asked to see the signatures. */
	readonly open = output<void>();

	protected readonly icons = SIGNATURE_ICONS;

	protected readonly view = computed<SignatureStatusView>(() =>
		signatureStatusView(this.summary(), this.translate.currentLang() ?? undefined),
	);
	protected readonly tone = computed(() => TONE_CLASS[this.view().tone]);
}
