/** Inherited public tokens use local fallback aliases so host themes are never shadowed. */
export const viewerStyles = `
:host {
  --_vv-bg:var(--vv-background,#fbfaf7); --_vv-surface:var(--vv-surface,#fff);
  --_vv-secondary:var(--vv-secondary,#f2efe8); --_vv-ink:var(--vv-ink,#1a1d21);
  --_vv-muted:var(--vv-muted,#5a626e); --_vv-border:var(--vv-border,#e6e2d9);
  --_vv-accent:var(--vv-accent,#c2431f); --_vv-accent-soft:var(--vv-accent-soft,rgba(194,67,31,.08));
  display:flex; flex-direction:column; min-width:0; min-height:320px; height:100%; overflow:hidden;
  color:var(--_vv-ink); background:var(--_vv-bg); font:12px/1.5 ui-sans-serif,system-ui,-apple-system,BlinkMacSystemFont,"Segoe UI",sans-serif; contain:layout;
}
* { box-sizing:border-box; }
[hidden] { display:none !important; }
button,select,input,textarea { font:inherit; color:inherit; border:1px solid var(--_vv-border); background:var(--_vv-surface); border-radius:4px; min-height:28px; padding:3px 8px; }
button { cursor:pointer; }
button:hover:enabled { background:var(--_vv-accent-soft); border-color:var(--_vv-accent); }
button:disabled { opacity:.4; cursor:default; }
button:focus-visible,select:focus-visible,input:focus-visible,textarea:focus-visible,summary:focus-visible,[tabindex]:focus-visible { outline:2px solid var(--vv-focus,var(--_vv-accent)); outline-offset:2px; }
button[aria-pressed="true"] { color:var(--_vv-accent); background:var(--_vv-accent-soft); }
.toolbar { flex:none; min-width:0; border-bottom:1px solid var(--_vv-border); background:color-mix(in srgb,var(--_vv-secondary) 55%,var(--_vv-surface)); }
.ribbon-primary { height:32px; display:flex; align-items:center; gap:6px; padding:2px 8px; }
.page-picker { display:flex; min-width:0; align-items:center; gap:8px; color:var(--_vv-muted); font-size:11px; }
.page-picker select { width:190px; max-width:42vw; height:26px; min-height:26px; padding:2px 6px; font-size:11px; color:var(--_vv-ink); }
.page-stepper { display:flex; align-items:center; gap:2px; }
.page-stepper button { width:26px; height:26px; min-height:26px; padding:0; font-size:18px; line-height:22px; background:transparent; border-color:transparent; }
.spacer { flex:1; }
.local-label { color:var(--_vv-muted); font-size:10px; white-space:nowrap; }
.ribbon-tabs { height:35px; display:flex; align-items:stretch; padding:0 4px; border-bottom:1px solid var(--_vv-border); }
.ribbon-tabs button { position:relative; min-height:34px; padding:7px 14px; border:0; border-radius:0; background:transparent; color:var(--_vv-muted); font-size:12px; font-weight:500; }
.ribbon-tabs button[aria-selected="true"] { color:var(--_vv-ink); }
.ribbon-tabs button[aria-selected="true"]::after { content:""; position:absolute; inset:auto 0 -1px; height:2.5px; background:var(--_vv-accent); }
.ribbon-content { min-height:82px; display:flex; align-items:stretch; gap:0; padding:2px 4px; overflow-x:auto; }
.ribbon-group { display:flex; flex-direction:column; justify-content:space-between; flex:none; gap:2px; padding:3px 10px 4px; border-right:1px solid var(--_vv-border); }
.group-label { font-size:9px; line-height:10px; color:var(--_vv-muted); text-align:center; }
.ribbon-actions { display:flex; align-items:stretch; gap:2px; flex:1; }
.ribbon-command { display:flex; flex-direction:column; justify-content:center; align-items:center; gap:2px; min-width:70px; padding:2px 6px; border-color:transparent; background:transparent; font-size:11px; white-space:nowrap; }
.command-icon { font:22px/26px Georgia,serif; color:var(--_vv-muted); }
.ribbon-command[data-chrome="edit"] .command-icon { font-family:ui-serif,Georgia,serif; text-decoration:underline; text-decoration-thickness:1px; text-underline-offset:3px; }
.ribbon-hint { align-self:center; flex:none; margin:0; padding:0 16px; font-size:11px; line-height:18px; color:var(--_vv-muted); }
.search-group { width:330px; }
.search-controls { display:flex; align-items:center; gap:4px; flex-wrap:wrap; }
.search-controls label { display:flex; align-items:center; min-width:0; gap:6px; flex:1; font-size:11px; }
.search-controls input { width:100%; min-width:0; padding:3px 6px; font-size:11px; }
.search-controls button { padding:3px 6px; font-size:10px; background:var(--_vv-secondary); }
.search-controls [role="status"] { display:block; width:100%; min-height:15px; font-size:10px; line-height:15px; color:var(--_vv-muted); overflow:hidden; white-space:nowrap; text-overflow:ellipsis; }
.workspace { position:relative; display:flex; flex:1; min-width:0; min-height:0; overflow:hidden; }
.page-rail { flex:0 0 180px; width:180px; min-width:0; overflow:auto; border-right:1px solid var(--_vv-border); background:color-mix(in srgb,var(--_vv-secondary) 40%,var(--_vv-bg)); }
.pane-heading { height:37px; display:flex; align-items:center; justify-content:space-between; gap:8px; padding:8px 10px; border-bottom:1px solid var(--_vv-border); font-size:11px; font-weight:600; }
.pane-heading [data-page-count],.inspector-kind { color:var(--_vv-muted); font-size:10px; font-weight:400; }
.page-list { list-style:none; padding:8px 5px; margin:0; display:flex; flex-direction:column; gap:4px; }
.page-link { display:flex; align-items:center; position:relative; width:100%; min-width:0; gap:4px; padding:4px 5px 4px 2px; border:0; border-radius:0; background:transparent; color:var(--_vv-muted); text-align:left; }
.page-link[aria-current="page"] { background:var(--_vv-accent-soft); }
.page-link[aria-current="page"]::before { content:""; position:absolute; left:0; top:4px; bottom:4px; width:3px; border-radius:0 2px 2px 0; background:var(--_vv-accent); }
.page-number { flex:0 0 20px; width:20px; padding-right:3px; font-size:10px; text-align:right; }
.page-card { display:flex; flex-direction:column; align-items:flex-start; min-width:0; width:100%; min-height:98px; gap:3px; padding:10px; border:1px solid var(--_vv-border); border-radius:2px; background:var(--_vv-surface); }
.page-link[aria-current="page"] .page-card { border-color:var(--_vv-accent); }
.page-link[aria-current="page"] .page-number { color:var(--_vv-accent); font-weight:600; }
.page-mark { color:var(--_vv-muted); font-size:20px; line-height:23px; }
.page-name { color:var(--_vv-ink); font-size:11px; font-weight:500; line-height:16px; overflow-wrap:anywhere; }
.page-meta { color:var(--_vv-muted); font-size:9px; line-height:14px; }
.page-empty { margin:12px; color:var(--_vv-muted); font-size:11px; }
.viewport { position:relative; display:flex; flex-direction:column; flex:1; overflow:auto; min-width:0; min-height:0; padding:16px 4px; background:var(--_vv-bg); }
.paper { display:block; flex:none; background:white; box-shadow:var(--vv-shadow,0 2px 12px rgb(0 0 0 / 14%)); margin:auto; }
.paper [data-shape-id] { cursor:pointer; }
.paper [data-selected="true"] > [data-geometry] { stroke:var(--_vv-accent) !important; stroke-width:.025 !important; }
.inspector-pane { flex:0 0 288px; width:288px; min-width:0; overflow:auto; border-left:1px solid var(--_vv-border); background:var(--_vv-bg); }
.inspector-body { display:flex; flex-direction:column; gap:12px; padding:12px 10px; }
.inspector-card,.edit-controls,.layer-controls { min-width:0; padding:8px; margin:0; border:1px solid var(--_vv-border); border-radius:4px; background:var(--_vv-surface); font-size:11px; line-height:17px; }
.document-card h2 { margin:0 0 8px; font-size:10px; font-weight:500; letter-spacing:.5px; text-transform:uppercase; color:var(--_vv-muted); }
.current-page-name { margin:0 0 8px; font-size:12px; font-weight:500; overflow-wrap:anywhere; }
.document-card dl { display:grid; grid-template-columns:1fr auto; gap:4px 10px; margin:0; font-size:11px; }
.document-card dt { color:var(--_vv-muted); }
.document-card dd { margin:0; text-align:right; overflow-wrap:anywhere; }
.selection-hint { margin:0; padding:0 2px; color:var(--_vv-muted); font-size:11px; }
summary { cursor:pointer; min-height:22px; line-height:22px; font-weight:500; }
summary::marker { color:var(--_vv-muted); font-size:10px; }
.shape-inspector div { max-height:260px; overflow:auto; }
.shape-inspector p { margin:6px 0; overflow-wrap:anywhere; }
.shape-inspector dl { display:grid; grid-template-columns:minmax(60px,1fr) 1.5fr; gap:6px 10px; }
.shape-inspector dt { font-weight:500; }
.shape-inspector dd { margin:0; overflow-wrap:anywhere; white-space:pre-wrap; }
.shape-inspector h3 { margin:12px 0 5px; font-size:11px; }
.shape-inspector a { color:var(--_vv-accent); overflow-wrap:anywhere; }
.edit-controls p,.layer-controls p { margin:6px 0; color:var(--_vv-muted); overflow-wrap:anywhere; }
.edit-controls textarea { display:block; width:100%; resize:vertical; max-height:180px; margin-top:4px; background:var(--_vv-secondary); }
.edit-controls label { display:block; margin-top:8px; }
.edit-actions { display:flex; flex-wrap:wrap; gap:4px; margin-top:8px; }
.edit-actions button { font-size:11px; padding:3px 7px; }
.edit-actions [data-edit="apply"]:enabled { background:var(--_vv-accent); border-color:var(--_vv-accent); color:var(--vv-accent-ink,#fff); }
.edit-controls [data-edit-error] { color:var(--vv-danger,#b42318); }
.edit-controls [data-edit-diagnostics] { max-height:100px; overflow:auto; padding-left:16px; }
.layer-controls [data-layer-list] { max-height:240px; overflow:auto; }
.layer-controls fieldset { border:1px solid var(--_vv-border); margin:6px 0; padding:6px; min-width:0; }
.layer-controls legend { max-width:100%; overflow-wrap:anywhere; }
.layer-controls fieldset div { display:flex; align-items:center; gap:6px; margin:3px 0; }
.layer-controls label { display:flex; align-items:center; gap:6px; overflow-wrap:anywhere; }
.layer-controls input { flex-shrink:0; width:16px; height:16px; min-height:16px; padding:0; accent-color:var(--_vv-accent); }
.notes ul { max-height:220px; overflow:auto; padding-left:17px; margin:6px 0 0; color:var(--_vv-muted); }
.notes li { margin:5px 0; overflow-wrap:anywhere; }
.notes-strip { flex:none; display:flex; align-items:center; justify-content:space-between; min-height:25.5px; padding:0 12px; border-top:1px solid var(--_vv-border); background:var(--_vv-bg); color:var(--_vv-muted); font-size:10px; }
.notes-strip button { min-height:24.5px; padding:2px 0; border:0; background:transparent; color:var(--_vv-muted); font-size:11px; }
.notes-strip button::before { content:"▸"; margin-right:6px; }
.notes-strip button[aria-expanded="true"]::before { content:"▾"; }
.status { flex:none; display:flex; align-items:center; justify-content:space-between; gap:12px; min-width:0; min-height:29px; padding:0 8px 0 12px; border-top:1px solid var(--_vv-border); color:var(--_vv-muted); background:var(--_vv-surface); font-size:10px; }
.status-message { display:flex; gap:16px; min-width:0; }
.status-message span { overflow:hidden; white-space:nowrap; text-overflow:ellipsis; }
.zoom-controls { display:flex; flex:none; align-items:center; gap:3px; }
.zoom-controls button { min-height:24px; height:24px; padding:1px 6px; border-color:transparent; background:transparent; font-size:10px; }
.zoom-controls [data-action="in"],.zoom-controls [data-action="out"] { min-width:24px; padding:0; font-size:17px; line-height:20px; }
.zoom-divider { width:1px; height:14px; background:var(--_vv-border); margin:0 3px; }
.zoom { min-width:40px; font-size:10px; text-align:center; font-variant-numeric:tabular-nums; }
.empty { max-width:360px; margin:60px auto; padding:16px; text-align:center; color:var(--_vv-muted); font-size:12px; }
.empty strong { display:block; color:var(--_vv-ink); font-size:18px; font-weight:500; margin-bottom:8px; }
@media(max-width:1100px) { .ribbon-hint { display:none; } [data-diagnostics] { display:none; } }
@media(max-width:760px) {
  button,select,input,textarea { min-height:44px; }
  button { min-width:44px; }
  summary { min-height:44px; line-height:44px; }
  .ribbon-primary { height:52px; padding:4px 8px; }
  .page-picker { flex:1; }
  .page-picker select { width:100%; max-width:none; min-height:44px; height:44px; font-size:12px; }
  .page-stepper button { width:44px; min-height:44px; height:44px; }
  .local-label,.ribbon-primary>.spacer { display:none; }
  .ribbon-tabs { height:44px; }
  .ribbon-tabs button { min-height:44px; padding:8px 16px; }
  .ribbon-content { min-height:86px; }
  .ribbon-group { padding:4px 8px; }
  .search-group { width:100%; border-right:0; }
  .search-controls { gap:4px; }
  .search-controls label { font-size:12px; }
  .search-controls input { font-size:12px; }
  .search-controls button { min-width:44px; font-size:0; padding:3px; }
  .search-controls [data-action="search-previous"]::before { content:"‹"; font-size:22px; }
  .search-controls [data-action="search-next"]::before { content:"›"; font-size:22px; }
  #home-panel>.ribbon-group:not(.search-group) { display:none; }
  .ribbon-command { min-width:90px; }
  .workspace { display:block; overflow:auto; }
  .page-rail { width:100%; max-height:190px; border-right:0; border-bottom:1px solid var(--_vv-border); }
  .page-list { flex-direction:row; overflow-x:auto; }
  .page-list li { flex:0 0 155px; }
  .page-card { min-height:86px; }
  .viewport { height:340px; min-height:260px; padding:16px 4px; }
  .inspector-pane { width:100%; overflow:visible; border-left:0; border-top:1px solid var(--_vv-border); }
  .inspector-body { gap:8px; padding:10px; }
  .inspector-card,.edit-controls,.layer-controls { font-size:12px; }
  .layer-controls label { min-height:44px; }
  .layer-controls input { min-width:20px; width:20px; height:20px; min-height:20px; }
  .notes-strip { min-height:44px; padding:0 10px; }
  .notes-strip button { min-height:44px; }
  .status { flex-wrap:wrap; gap:0; padding:4px 8px; }
  .status-message { width:100%; min-height:18px; }
  .zoom-controls { justify-content:flex-end; width:100%; }
  .zoom-controls button { min-height:44px; height:44px; min-width:44px; font-size:11px; }
  .zoom-controls [data-action="in"],.zoom-controls [data-action="out"] { min-width:44px; font-size:20px; }
  .zoom { font-size:11px; }
}
@media(pointer:coarse) {
  button,select,input,textarea { min-height:44px; }
  button { min-width:44px; }
  summary,.layer-controls label { min-height:44px; line-height:44px; }
  .ribbon-primary { min-height:52px; height:52px; }
  .page-picker select,.page-stepper button { min-height:44px; height:44px; }
  .page-stepper button { min-width:44px; }
  .ribbon-tabs { min-height:44px; height:44px; }
  .ribbon-tabs button { min-height:44px; }
  .ribbon-content { min-height:96px; }
  .search-group { width:380px; max-width:100%; }
  .zoom-controls button,.notes-strip button { min-height:44px; height:44px; }
  .zoom-controls [data-action="in"],.zoom-controls [data-action="out"] { min-width:44px; }
  .layer-controls input { min-height:20px; }
}
@media(prefers-reduced-motion:reduce) { * { scroll-behavior:auto !important; } }
`;
