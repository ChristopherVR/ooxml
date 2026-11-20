/**
 * The signature badge row under the Ribbon (`ui/signature-badge.ts`). The tone
 * comes from `signatureStatusView`; phones and slide shows hide the row, as
 * the Vue chrome hides its badge on the mobile layout.
 */
export const SIGNATURE_BADGE_CSS = `
.pptxv-signature-badge-row {
	display: flex;
	align-items: center;
	padding: 4px 12px;
}
.pptxv-signature-badge-row[hidden] { display: none; }
.pptxv-signature-badge {
	--pptxv-signature-tone: var(--pptx-muted-foreground);
	display: inline-flex;
	align-items: center;
	gap: 4px;
	padding: 2px 8px;
	border: 1px solid color-mix(in srgb, var(--pptxv-signature-tone) 40%, transparent);
	border-radius: 6px;
	background: color-mix(in srgb, var(--pptxv-signature-tone) 14%, transparent);
	color: var(--pptxv-signature-tone);
	font: inherit;
	font-size: 11px;
	font-weight: 500;
	cursor: pointer;
}
.pptxv-signature-badge:hover { background: color-mix(in srgb, var(--pptxv-signature-tone) 24%, transparent); }
.pptxv-signature-badge:focus-visible { outline: 2px solid var(--pptx-ring); outline-offset: 1px; }
.pptxv-signature-badge[data-tone='success'] { --pptxv-signature-tone: #16a34a; }
.pptxv-signature-badge[data-tone='caution'] { --pptxv-signature-tone: #d97706; }
.pptxv-signature-badge[data-tone='danger'] { --pptxv-signature-tone: #dc2626; }
.pptxv-signature-badge svg { width: 14px; height: 14px; flex: none; }
.pptxv.pptxv-presenting .pptxv-signature-badge-row { display: none; }
@media (max-width: 767px), (max-width: 1023px) and (max-height: 520px) {
	.pptxv-signature-badge-row { display: none; }
}
`;
