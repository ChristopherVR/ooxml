/** Inherited public tokens use local fallback aliases so host themes are never shadowed. */
export const viewerStyles = `
:host {
  --_vv-bg:var(--vv-background,#eceae8); --_vv-surface:var(--vv-surface,#fff);
  --_vv-secondary:var(--vv-secondary,#f3f2f1); --_vv-ink:var(--vv-ink,#1f1f1f);
  --_vv-muted:var(--vv-muted,#605e5c); --_vv-border:var(--vv-border,#e1dfdd);
  --_vv-accent:var(--vv-accent,#3955a3); --_vv-accent-soft:var(--vv-accent-soft,rgba(57,85,163,.10));
  display:flex; flex-direction:column; min-width:0; min-height:320px; height:100%; overflow:hidden;
  color:var(--_vv-ink); background:var(--_vv-bg); font:12px/1.5 "Segoe UI",Arial,sans-serif; contain:layout;
}
* { box-sizing:border-box; }
[hidden] { display:none !important; }
button,select,input,textarea { font:inherit; color:inherit; border:1px solid var(--_vv-border); background:var(--_vv-surface); border-radius:4px; min-height:28px; padding:3px 8px; }
button { cursor:pointer; }
button:hover:enabled { background:var(--_vv-accent-soft); border-color:var(--_vv-accent); }
button:disabled { opacity:.4; cursor:default; }
button:focus-visible,select:focus-visible,input:focus-visible,textarea:focus-visible,summary:focus-visible,[tabindex]:focus-visible { outline:2px solid var(--vv-focus,var(--_vv-accent)); outline-offset:2px; }
button[aria-pressed="true"] { color:var(--_vv-accent); background:var(--_vv-accent-soft); }
.toolbar { flex:none; min-width:0; border-bottom:1px solid var(--_vv-border); background:var(--_vv-surface); }
.spacer { flex:1; }
.ribbon-content { min-height:82px; display:flex; align-items:stretch; gap:0; padding:2px 4px; overflow-x:auto; }
.ribbon-group { display:flex; flex-direction:column; justify-content:space-between; flex:none; gap:2px; padding:3px 10px 4px; border-right:1px solid var(--_vv-border); }
.group-label { font-size:10px; line-height:12px; color:var(--_vv-muted); text-align:center; }
.ribbon-actions { display:flex; align-items:stretch; gap:2px; flex:1; }
.ribbon-command { display:flex; flex-direction:column; justify-content:center; align-items:center; gap:2px; min-width:70px; padding:2px 6px; border-color:transparent; background:transparent; font-size:11px; white-space:nowrap; }
.ribbon-hint { align-self:center; flex:none; margin:0; padding:0 16px; font-size:11px; line-height:18px; color:var(--_vv-muted); }
.workspace { position:relative; display:flex; flex:1; min-width:0; min-height:0; overflow:hidden; }
.pane-heading { height:37px; display:flex; align-items:center; justify-content:space-between; gap:8px; padding:8px 10px; border-bottom:1px solid var(--_vv-border); font-size:11px; font-weight:600; }
.inspector-kind { color:var(--_vv-muted); font-size:10px; font-weight:400; }
.viewport { position:relative; display:flex; flex-direction:column; flex:1; overflow:auto; min-width:0; min-height:0; padding:16px 4px; background:var(--_vv-bg); }
.paper { display:block; flex:none; background:white; box-shadow:var(--vv-shadow,0 2px 8px rgb(0 0 0 / 14%)); margin:auto; }
.paper [data-shape-id] { cursor:pointer; }
.paper [data-selected="true"] > [data-geometry] { stroke:var(--_vv-accent) !important; stroke-width:.025 !important; }
.inspector-pane { flex:0 0 288px; width:288px; min-width:0; overflow:auto; border-left:1px solid var(--_vv-border); background:var(--_vv-secondary); }
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
.edit-controls [data-geometry] { min-width:0; margin:8px 0 0; padding:6px; border:1px solid var(--_vv-border); }
.edit-controls [data-geometry-field] { display:block; width:100%; min-width:0; margin-top:4px; }
.edit-controls [data-geometry-error] { color:var(--vv-danger,#b42318); }
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
.empty { max-width:360px; margin:60px auto; padding:16px; text-align:center; color:var(--_vv-muted); font-size:12px; }
.empty strong { display:block; color:var(--_vv-ink); font-size:18px; font-weight:500; margin-bottom:8px; }
@media(max-width:1100px) { .ribbon-hint { display:none; } [data-diagnostics] { display:none; } }
@media(max-width:760px) {
  button,select,input,textarea { min-height:44px; }
  button { min-width:44px; }
  summary { min-height:44px; line-height:44px; }
  .ribbon-content { min-height:86px; }
  .ribbon-group { padding:4px 8px; }
  #home-panel>.ribbon-group:not(.search-group):not(.undo-group) { display:none; }
  .ribbon-command { min-width:90px; }
  .workspace { display:block; overflow:auto; }
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
}
@media(pointer:coarse) {
  button,select,input,textarea { min-height:44px; }
  button { min-width:44px; }
  summary,.layer-controls label { min-height:44px; line-height:44px; }
  .ribbon-content { min-height:96px; }
  .notes-strip button { min-height:44px; height:44px; }
  .layer-controls input { min-height:20px; }
}
/* Suite ribbon: tabs lead directly into Visio command groups. */
.ribbon-content { min-height:104px; padding:4px 4px; }
.ribbon-hint { order:3; padding:10px 18px; }
.ribbon-command { min-width:82px; }
.pane-close { display:none; }
@media(max-width:760px) {
 .ribbon-content { min-height:54px; padding:4px 0; }
 .search-feedback,.group-label { display:none; }
 .ribbon-content { overflow-x:auto; scroll-snap-type:x mandatory; }
 .workspace { display:flex; position:relative; overflow:hidden; }
 .viewport { flex:1; width:100%; height:100%; min-height:0; padding:16px 4px; }
 .inspector-pane { position:absolute; inset:0 0 0 auto; width:min(320px,100%); z-index:4; overflow:auto; border-left:1px solid var(--_vv-border); box-shadow:-8px 0 24px #0002; }
 .pane-close { display:block; width:44px; height:44px; padding:0; font-size:22px; background:transparent; border:0; }
 .inspector-kind { margin-left:auto; }
 .notes-strip { display:none; }
 .status { min-height:48px; padding:2px 8px; }
 .status-message { display:none; }
}
.toolbar { position:relative; }
.ribbon-tools { display:flex; flex:none; order:0; }
.tools-content { display:flex; align-items:stretch; }
.ribbon-tools>summary { display:none; }
.ribbon-command { font-size:12px; min-width:88px; }
.ribbon-hint { font-size:12px; line-height:20px; flex:1; }
.group-label { font-size:11px; line-height:14px; }
.inspector-card,.edit-controls,.layer-controls { font-size:12px; line-height:18px; }
.selection-hint,.document-card dl { font-size:12px; }
.pane-heading { font-size:12px; }
.inspector-pane .pane-heading { position:sticky; top:0; z-index:1; background:var(--_vv-secondary); flex:none; }
.status .notes-strip { border:0; padding:0 8px; background:transparent; min-height:28px; }
.status .notes-strip>span { display:none; }
.status-message { flex:1; }
@media(max-width:760px) {
 #home-panel { overflow:visible; min-height:54px; padding:4px; }
 .ribbon-tools { display:block; flex:0 0 84px; width:84px; order:1; }
 .ribbon-tools>summary { display:flex; min-height:44px; align-items:center; justify-content:center; gap:6px; padding:0 10px; border:1px solid var(--_vv-border); background:var(--_vv-surface); border-radius:4px; font-size:12px; cursor:pointer; list-style:none; }
 .ribbon-tools>summary::-webkit-details-marker { display:none; }
 .ribbon-tools .tools-content { display:flex; position:absolute; top:100%; left:4px; right:4px; z-index:6; flex-direction:column; padding:12px; gap:12px; border:1px solid var(--_vv-border); border-radius:4px; background:var(--_vv-surface); box-shadow:0 12px 28px #0003; max-height:60vh; overflow:auto; }
 .ribbon-tools:not([open])>.tools-content { display:none; }
 .ribbon-tools .ribbon-group { display:flex; width:100%; min-height:64px; border:0; padding:0; }
 .ribbon-tools .ribbon-actions { justify-content:space-between; }
 .ribbon-tools .ribbon-command { min-width:88px; font-size:12px; }
 .ribbon-tools .search-feedback { display:block; }
 .status { flex-wrap:nowrap; gap:4px; min-height:48px; }
 .status .notes-strip { display:none; }
}
@media(prefers-reduced-motion:reduce) { * { scroll-behavior:auto !important; } }
`;
