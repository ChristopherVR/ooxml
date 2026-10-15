// Page Layout view CSS (shadow DOM). Paper stays white and the surround grey like Excel; chrome
// colours come from the editor's `--xve-*` tokens.
export const PAGE_LAYOUT_CSS = `
.xve-grid-host[data-view="pageLayout"]>.xg-root{display:none}
.xpl-root{position:absolute;inset:0;overflow:auto;background:var(--xve-page-surround,#e6e6e6);
	font-family:"Segoe UI",system-ui,sans-serif;outline:none;color:#000;-webkit-user-select:none;user-select:none}
.xpl-root[hidden]{display:none}
.xpl-sizer{display:grid;grid-template-columns:20px 1fr;grid-template-rows:20px 1fr;position:relative}
.xpl-corner{position:sticky;left:0;top:0;z-index:7;background:var(--xve-header-bg,#f3f3f3);border-right:1px solid #c6c6c6;border-bottom:1px solid #c6c6c6}
.xpl-ruler-h{position:sticky;top:0;z-index:6;height:20px;background:var(--xve-header-bg,#f3f3f3);overflow:hidden;border-bottom:1px solid #c6c6c6}
.xpl-ruler-v{position:sticky;left:0;z-index:6;width:20px;background:var(--xve-header-bg,#f3f3f3);overflow:hidden;border-right:1px solid #c6c6c6}
.xpl-ruler-area{position:absolute;background:#fff}
.xpl-ruler-margin{position:absolute;background:#c8c8c8}
.xpl-ruler-h .xpl-ruler-area,.xpl-ruler-h .xpl-ruler-margin{top:0;bottom:0}
.xpl-ruler-v .xpl-ruler-area,.xpl-ruler-v .xpl-ruler-margin{left:0;right:0}
.xpl-tick{position:absolute;background:#767676}
.xpl-ruler-h .xpl-tick{bottom:0;width:1px}
.xpl-ruler-v .xpl-tick{right:0;height:1px}
.xpl-ruler-num{position:absolute;font-size:10px;line-height:1;color:#444}
.xpl-ruler-v .xpl-ruler-num{left:2px}
.xpl-canvas{position:relative;overflow:hidden}
.xpl-scaled{position:absolute;left:0;top:0;transform-origin:0 0}
.xpl-page{position:absolute;box-sizing:border-box;background:#fff;box-shadow:0 1px 5px rgba(0,0,0,.3)}
.xpl-content{position:absolute;overflow:hidden}
.xpl-inner{position:absolute;left:0;top:0;transform-origin:0 0;font-family:Calibri,sans-serif}
.xpl-inner table{border-collapse:collapse;table-layout:fixed}
.xpl-inner td{overflow:hidden;padding:0 3px;box-sizing:border-box;line-height:1.2}
.xpl-colhdr,.xpl-rowhdr{position:absolute;box-sizing:border-box;display:flex;align-items:center;justify-content:center;
	font-size:11px;color:var(--xve-header-fg,#444);background:var(--xve-header-bg,#f3f3f3);border:1px solid var(--xve-header-line,#d4d4d4);overflow:hidden}
.xpl-blank{position:absolute;inset:0;pointer-events:none;z-index:1;display:flex;align-items:center;justify-content:center;color:#767676;font-size:14px;cursor:cell}
.xpl-band{position:absolute;display:flex;line-height:1.2;box-sizing:border-box;cursor:text}
.xpl-section{flex:1 1 0;min-width:0;white-space:pre-wrap;overflow:hidden;padding:0 2px;border:1px solid transparent}
.xpl-band:hover .xpl-section{border-color:#bbb;border-style:dashed}
.xpl-placeholder{color:#8a8a8a;text-align:center!important}
.xpl-band.xpl-editing{background:#fff;z-index:4;box-shadow:0 0 0 1px #217346}
.xpl-band-input{flex:1 1 0;min-width:0;box-sizing:border-box;font:inherit;font-size:13px;border:1px solid #bbb;padding:1px 3px;margin:0}
.xpl-sel{position:absolute;left:0;top:0;pointer-events:none;z-index:3}
.xpl-shade{position:absolute;background:var(--xve-selection,rgba(33,115,70,.12))}
.xpl-active{position:absolute;box-sizing:border-box;border:2px solid var(--xve-selection-border,#217346)}
.xpl-edit{position:absolute;box-sizing:border-box;z-index:5;border:2px solid var(--xve-selection-border,#217346);
	background:#fff;font-family:Calibri,sans-serif;padding:0 3px;margin:0;outline:none}
.xpl-error{padding:24px;color:#a4262c}
`;
