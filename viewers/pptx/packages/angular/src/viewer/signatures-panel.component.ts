/**
 * signatures-panel.component.ts: Read-only digital-signature status panel.
 *
 * Selector: `pptx-signatures-panel`
 *
 * Shows the deck's verified signature state and lists each signature (signer,
 * its own state, issuer, signing time). A deck reads as signed only when its
 * signatures verify: the wording, tone and icon all come from the shared
 * `signatureStatusView` (ooxml-ui/pptx) in the viewer's current language, so
 * this panel keeps no status logic of its own. Purely presentational: it
 * inspects, it never signs or strips.
 *
 * Usage:
 * ```html
 * <pptx-signatures-panel [summary]="loader.signatureSummary()" />
 * ```
 */

import { ChangeDetectionStrategy, Component, computed, inject, input } from '@angular/core';
import { LucideDynamicIcon } from '@lucide/angular';
import { TranslatePipe, TranslateService } from '@ngx-translate/core';
import { signatureStatusView } from 'ooxml-ui/pptx';
import type { SignatureStatusView } from 'ooxml-ui/pptx';
import type { SignatureSummary } from 'pptx-viewer-core';

import { SIGNATURE_ICONS } from './signature-icons';

@Component({
	selector: 'pptx-signatures-panel',
	standalone: true,
	changeDetection: ChangeDetectionStrategy.OnPush,
	imports: [TranslatePipe, LucideDynamicIcon],
	template: `
		<section
			class="pptx-ng-signatures"
			[attr.aria-label]="'pptx.digitalSignatures.ariaLabel' | translate"
			[attr.data-pptx-signature-state]="view().state"
		>
			<header
				class="pptx-ng-signatures__header"
				[class]="'pptx-ng-signatures__header--' + view().tone"
			>
				<svg
					class="pptx-ng-signatures__icon"
					aria-hidden="true"
					[lucideIcon]="icons[view().icon]"
				></svg>
				<span class="pptx-ng-signatures__title">{{
					view().present ? view().label : view().heading
				}}</span>
			</header>

			@if (!view().present) {
				<p class="pptx-ng-signatures__empty">
					{{ 'pptx.digitalSignatures.noSignatures' | translate }}
				</p>
			} @else {
				<p class="pptx-ng-signatures__description">{{ view().description }}</p>
				<ul class="pptx-ng-signatures__list">
					@for (row of view().rows; track row.path) {
						<li
							class="pptx-ng-signatures__item"
							[class]="'pptx-ng-signatures__item--' + row.tone"
							[attr.data-pptx-signature-row-state]="row.state"
						>
							<div class="pptx-ng-signatures__item-main">
								<span class="pptx-ng-signatures__signer" [attr.title]="row.subject ?? null">{{
									row.signer
								}}</span>
								<span
									class="pptx-ng-signatures__badge"
									[class]="'pptx-ng-signatures__badge--' + row.tone"
								>
									{{ row.label }}
								</span>
							</div>

							<dl class="pptx-ng-signatures__meta">
								@if (row.issuer; as issuer) {
									<dt>{{ 'pptx.digitalSignatures.issuer' | translate }}</dt>
									<dd>{{ issuer }}</dd>
								}
								@if (row.signedAt; as signedAt) {
									<dt>{{ view().signedAtCaption }}</dt>
									<dd>{{ signedAt }}</dd>
								}
							</dl>
						</li>
					}
				</ul>
				<p class="pptx-ng-signatures__note">{{ view().removedOnSave }}</p>
			}
		</section>
	`,
	styles: [
		`
			.pptx-ng-signatures {
				font-family: system-ui, sans-serif;
				font-size: 13px;
				color: #1f2937;
				background: #fff;
				border: 1px solid #e5e7eb;
				border-radius: 8px;
				overflow: hidden;
			}

			.pptx-ng-signatures__header {
				display: flex;
				align-items: center;
				gap: 8px;
				padding: 10px 12px;
				font-weight: 600;
				border-bottom: 1px solid #e5e7eb;
			}

			.pptx-ng-signatures__header--success {
				background: #ecfdf5;
				color: #065f46;
			}

			.pptx-ng-signatures__header--caution {
				background: #fffbeb;
				color: #92400e;
			}

			.pptx-ng-signatures__header--danger {
				background: #fef2f2;
				color: #991b1b;
			}

			.pptx-ng-signatures__header--neutral {
				background: #f9fafb;
				color: #374151;
			}

			.pptx-ng-signatures__icon {
				width: 16px;
				height: 16px;
				flex: none;
			}

			.pptx-ng-signatures__title {
				flex: 1;
			}

			.pptx-ng-signatures__empty,
			.pptx-ng-signatures__description,
			.pptx-ng-signatures__note {
				margin: 0;
				padding: 10px 12px;
				color: #4b5563;
			}

			.pptx-ng-signatures__empty {
				padding: 14px 12px;
				color: #6b7280;
			}

			.pptx-ng-signatures__note {
				font-size: 12px;
				color: #6b7280;
				border-top: 1px solid #f3f4f6;
			}

			.pptx-ng-signatures__list {
				list-style: none;
				margin: 0;
				padding: 0;
				border-top: 1px solid #f3f4f6;
			}

			.pptx-ng-signatures__item {
				padding: 10px 12px;
				border-bottom: 1px solid #f3f4f6;
				border-left: 3px solid transparent;
			}

			.pptx-ng-signatures__item:last-child {
				border-bottom: none;
			}

			.pptx-ng-signatures__item--success {
				border-left-color: #10b981;
			}

			.pptx-ng-signatures__item--caution {
				border-left-color: #f59e0b;
			}

			.pptx-ng-signatures__item--danger {
				border-left-color: #ef4444;
			}

			.pptx-ng-signatures__item--neutral {
				border-left-color: #9ca3af;
			}

			.pptx-ng-signatures__item-main {
				display: flex;
				align-items: center;
				gap: 8px;
				justify-content: space-between;
			}

			.pptx-ng-signatures__signer {
				font-weight: 600;
				word-break: break-word;
			}

			.pptx-ng-signatures__badge {
				flex: none;
				font-size: 11px;
				font-weight: 600;
				padding: 2px 8px;
				border-radius: 999px;
				white-space: nowrap;
			}

			.pptx-ng-signatures__badge--success {
				background: #d1fae5;
				color: #065f46;
			}

			.pptx-ng-signatures__badge--caution {
				background: #fef3c7;
				color: #92400e;
			}

			.pptx-ng-signatures__badge--danger {
				background: #fee2e2;
				color: #991b1b;
			}

			.pptx-ng-signatures__badge--neutral {
				background: #f3f4f6;
				color: #374151;
			}

			.pptx-ng-signatures__meta {
				display: grid;
				grid-template-columns: auto 1fr;
				gap: 2px 10px;
				margin: 6px 0 0;
				font-size: 12px;
				color: #4b5563;
			}

			.pptx-ng-signatures__meta dt {
				font-weight: 500;
				color: #6b7280;
			}

			.pptx-ng-signatures__meta dd {
				margin: 0;
				word-break: break-word;
			}
		`,
	],
})
export class SignaturesPanelComponent {
	private readonly translate = inject(TranslateService);

	/** The loaded deck's signatures (`PptxData.signatures`); absent for an unsigned deck. */
	readonly summary = input<SignatureSummary | undefined>(undefined);

	protected readonly icons = SIGNATURE_ICONS;

	/** What the panel shows, in the viewer's current language. */
	protected readonly view = computed<SignatureStatusView>(() =>
		signatureStatusView(this.summary(), this.translate.currentLang() ?? undefined),
	);
}
