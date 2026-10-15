import { attachScopedRibbonStyles } from './ribbon-scoped-styles';

const RIBBON_INSERT_STYLES = `
:host { display:inline-flex; align-items:stretch; flex:none; color:var(--pptx-foreground,#f8fafc); font:inherit; }
[hidden] { display:none !important; }
:host[hidden] { display:none !important; }
.stack { display:flex; flex-direction:column; justify-content:flex-start; gap:2px; }
.trigger,.list button { font:inherit; color:inherit; box-sizing:border-box; }
.trigger { display:inline-flex; align-items:center; justify-content:flex-start; gap:6px; min-height:24px; min-width:0; padding:0 6px; border:0; border-radius:4px; background:transparent; cursor:pointer; font-size:12px; line-height:16px; }
.trigger:hover:not(:disabled),.list button:hover { background:var(--pptx-accent,#33334d); }
.trigger:active:not(:disabled) { background:color-mix(in srgb,var(--pptx-accent,#33334d) 70%,var(--pptx-primary,#6366f1)); }
.trigger:disabled,.list button:disabled { opacity:.35; cursor:not-allowed; }
.trigger:focus-visible,.list button:focus-visible { outline:2px solid var(--pptx-ring,#6366f1); outline-offset:2px; }
svg { width:16px; height:16px; flex:none; fill:none; stroke:currentColor; stroke-width:1.5; stroke-linecap:round; stroke-linejoin:round; }
.trigger svg:first-child { color:var(--pptx-primary,#6366f1); }
.trigger svg:last-child { width:10px; height:10px; margin-inline-start:auto; }
/* Office large command: a 32px glyph over the label with the chevron beneath. */
.trigger.large { flex-direction:column; align-items:center; justify-content:flex-start; gap:2px; min-width:48px; max-width:80px; height:66px; padding:3px 4px; text-align:center; line-height:15px; }
.trigger.large svg:first-child { width:32px; height:32px; stroke-width:1.4; }
.trigger.large svg:last-child { width:8px; height:8px; margin:0; stroke-width:1.8; }
.menu { position:relative; display:inline-flex; }
/* Office split button: the face runs the default command, the chevron below opens the menu. */
.menu.split { flex-direction:column; align-items:stretch; }
.primary { display:flex; flex-direction:column; align-items:center; justify-content:flex-start; gap:2px; box-sizing:border-box; min-width:48px; max-width:80px; height:50px; padding:3px 4px 0; border:0; border-radius:4px 4px 0 0; background:transparent; color:inherit; font:inherit; font-size:12px; line-height:15px; text-align:center; cursor:pointer; }
.primary svg { width:32px; height:32px; flex:none; fill:none; stroke:var(--pptx-primary,#6366f1); stroke-width:1.4; stroke-linecap:round; stroke-linejoin:round; }
.primary:hover:not(:disabled) { background:var(--pptx-accent,#33334d); }
.primary:disabled { opacity:.35; cursor:not-allowed; }
.primary:focus-visible { outline:2px solid var(--pptx-ring,#6366f1); outline-offset:2px; }
.trigger.caret { justify-content:center; width:100%; min-height:0; height:16px; padding:0; border-radius:0 0 4px 4px; }
.trigger.caret svg:first-child { width:8px; height:8px; margin:0; color:inherit; stroke-width:1.8; }
.list .heading { padding:4px 12px 2px; font-size:11px; opacity:.7; }
.list.grid .heading { grid-column:1 / -1; padding:2px 2px 0; }
.list button[aria-current] { font-weight:600; background:var(--pptx-accent,#33334d); }
.list { position:fixed; z-index:1200; display:flex; flex-direction:column; min-width:170px; max-width:calc(100vw - 16px); max-height:65vh; overflow:auto; padding:4px 0; background:var(--pptx-popover,#111827); border:1px solid var(--pptx-border,#374151); border-radius:6px; box-shadow:0 8px 24px #0005; }
.list[hidden] { display:none; }
.list button { display:flex; align-items:center; gap:8px; width:100%; min-height:28px; padding:4px 12px; border:0; background:transparent; font-size:12px; text-align:left; cursor:pointer; }
.list svg { width:16px; height:16px; }
.list.grid { display:grid; grid-template-columns:repeat(8,32px); gap:2px; min-width:0; padding:6px; }
.list.grid button { width:32px; min-height:32px; height:32px; padding:0; justify-content:center; border-radius:4px; }
.list.grid svg { width:18px; height:18px; }
@media (max-width:767px) {
	:host { flex-wrap:wrap; max-width:100%; }
}
@media (pointer:coarse),(max-width:900px) {
	.list button { min-height:44px; }
	.list.grid { grid-template-columns:repeat(6,44px); }
	.list.grid button { width:44px; height:44px; }
	.trigger { min-height:44px; min-width:44px; }
	.trigger.caret { min-height:24px; height:24px; }
}
@media (forced-colors:active) {
	.list { border-color:ButtonText; }
	.trigger svg:first-child { color:ButtonText; }
	.trigger:focus-visible,.list button:focus-visible { outline-color:Highlight; }
	.trigger:disabled,.primary:disabled { color:GrayText; opacity:1; }
	.primary svg { stroke:ButtonText; }
}
`;

export function attachRibbonInsertStyles(doc: Document): void {
	attachScopedRibbonStyles(
		doc,
		'pptx-ui-insert-styles',
		'pptx-ui-ribbon-insert',
		RIBBON_INSERT_STYLES,
	);
}
