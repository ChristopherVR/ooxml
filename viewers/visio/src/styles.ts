export const viewerStyles = `
:host { --vv-accent:#126a64; --vv-ink:#182a33; --vv-muted:#5c6d76; --vv-border:#dce5e7; display:flex; flex-direction:column; min-height:320px; height:100%; color:var(--vv-ink); font:14px/1.5 system-ui,sans-serif; contain:layout; }
* { box-sizing:border-box; }
button,select,input,textarea { font:inherit; color:inherit; border:1px solid var(--vv-border); background:#fff; border-radius:7px; min-height:36px; padding:5px 10px; }
button { cursor:pointer; }
button:hover:enabled { background:#eef6f5; border-color:#8daca9; }
button:disabled { opacity:.45; cursor:default; }
button:focus-visible,select:focus-visible,input:focus-visible,textarea:focus-visible,[tabindex]:focus-visible { outline:3px solid #edb852; outline-offset:2px; }
.toolbar { display:flex; align-items:center; flex-wrap:wrap; gap:8px; padding:10px 14px; border-bottom:1px solid var(--vv-border); background:#fff; }
.toolbar[hidden] { display:none; }
.toolbar label { color:var(--vv-muted); }
.toolbar select { max-width:220px; }
.search-controls { display:flex; align-items:center; flex-wrap:wrap; gap:8px; flex-basis:100%; }
.search-controls label { display:flex; align-items:center; gap:6px; }
.search-controls input { width:210px; max-width:100%; }
.search-controls [role="status"] { color:var(--vv-muted); font-size:12px; }

.layer-controls { flex-basis:100%; font-size:12px; }
.layer-controls[hidden] { display:none; }
.layer-controls summary { cursor:pointer; min-height:28px; line-height:28px; }
.layer-controls p { margin:4px 0; color:var(--vv-muted); }
.layer-controls [data-layer-list] { max-height:240px; overflow:auto; }
.layer-controls fieldset { border:1px solid var(--vv-border); margin:6px 0; min-width:0; }
.layer-controls fieldset div { display:flex; align-items:center; gap:8px; margin:3px 0; }
.layer-controls label { display:flex; align-items:center; gap:8px; overflow-wrap:anywhere; }
.layer-controls input { flex-shrink:0; width:18px; min-height:18px; padding:0; }
.edit-controls { flex-basis:100%; min-width:0; font-size:12px; }
.edit-controls summary { cursor:pointer; min-height:28px; line-height:28px; }
.edit-controls p { margin:4px 0; overflow-wrap:anywhere; }
.edit-controls textarea { display:block; width:100%; resize:vertical; max-height:180px; }
.edit-controls label { display:block; }
.edit-actions { display:flex; flex-wrap:wrap; gap:8px; margin-top:8px; }
.edit-controls [data-edit-error] { color:#9e2c22; }
.edit-controls [data-edit-diagnostics] { max-height:100px; overflow:auto; }
.spacer { flex:1; }
.zoom { min-width:4em; text-align:center; font-variant-numeric:tabular-nums; }
.viewport { flex:1; overflow:auto; min-height:0; padding:32px; background-color:#eaf0f1; background-image:radial-gradient(#b8c8cb 0.7px,transparent 0.7px); background-size:16px 16px; }
.paper { display:block; background:white; box-shadow:0 3px 18px #18333b20; margin:0 auto; }
.paper [data-shape-id] { cursor:pointer; }
.paper [data-selected="true"] > [data-geometry] { stroke:#2677e9 !important; stroke-width:.025 !important; }
.status { padding:7px 14px; border-top:1px solid var(--vv-border); font-size:12px; background:#fff; color:var(--vv-muted); display:flex; gap:12px; justify-content:space-between; }
.shape-inspector { padding:8px 14px; font-size:12px; background:#fff; border-top:1px solid var(--vv-border); }
.shape-inspector summary { cursor:pointer; min-height:28px; line-height:28px; }
.shape-inspector div { max-height:200px; overflow:auto; }
.shape-inspector dl { display:grid; grid-template-columns:minmax(90px,1fr) 2fr; gap:6px 12px; }
.shape-inspector dt { font-weight:600; }
.shape-inspector dd { margin:0; overflow-wrap:anywhere; white-space:pre-wrap; }
.shape-inspector h3 { margin:12px 0 5px; font-size:12px; }
.shape-inspector a { color:var(--vv-accent); overflow-wrap:anywhere; }
.notes { font-size:12px; background:#fffaf0; border-top:1px solid var(--vv-border); padding:8px 14px; }
.notes summary { cursor:pointer; min-height:28px; line-height:28px; }
.notes ul { max-height:180px; overflow:auto; padding-left:18px; }
.notes li { margin:5px 0; }
.empty { max-width:400px; margin:60px auto; text-align:center; color:var(--vv-muted); }
.empty strong { display:block; color:var(--vv-ink); font-size:20px; margin-bottom:8px; }
@media(max-width:600px) { button,select,input,textarea { min-height:44px; min-width:44px; } .notes summary,.shape-inspector summary,.layer-controls summary,.edit-controls summary { min-height:44px; line-height:44px; } .viewport { padding:16px; } .toolbar { padding:8px; gap:5px; } .toolbar select { max-width:120px; } }
`;
