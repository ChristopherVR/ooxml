/**
 * PowerPoint's touches on the shared tab row (`pptx-ui-ribbon-tabs`). File is one of the tabs, in
 * the accent colour and filled while its backstage shows, so the shared File button never draws.
 * A selected tab reads in the foreground colour over the accent underline; a contextual tab keeps
 * its tint. Every value is an `--office-*` token the pptx bridge feeds from the viewer theme.
 */
export const RIBBON_TABS_CSS = `
.file { display: none; }
.head { min-height: var(--office-ribbon-tab-height); }
[role='tab'] { padding: 8px 14px; font-weight: 500; line-height: 18px; white-space: nowrap; }
[role='tab'][aria-selected='true']:not([data-contextual]) { color: var(--office-foreground); }
[role='tab'][data-tab='file'] { color: var(--office-accent); }
[role='tab'][data-tab='file'][aria-selected='true'] {
	background: var(--office-accent); color: var(--office-accent-foreground);
}
[role='tab'][data-tab='file'][aria-selected='true']::after { content: none; }
@media (min-width: 768px) and (max-width: 1000px) {
	.head { flex-wrap: wrap; }
	[role='tablist'] { flex: 0 0 auto; max-width: 100%; }
	[role='tab'] { padding-inline: 6px; }
}
@media (forced-colors: active) {
	[role='tab'][data-tab='file'][aria-selected='true'] {
		forced-color-adjust: none; background: Highlight; color: HighlightText;
	}
}
`;
