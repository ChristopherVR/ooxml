// Grid CSS (shadow DOM). UI colours inherit canonical Office tokens with legacy --xve-* overrides.
// Document cells keep their own colours; the sheet paper defaults to white.
import { MIN_TOUCH_TARGET_PX } from '../../touch-targets';
export const GRID_CSS = `
.xg-root [hidden]{display:none!important}
.xg-root{position:relative;overflow:hidden;width:100%;height:100%;min-height:0;outline:none;
	background:var(--xve-sheet-bg, var(--office-paper, #fff));color:#000;font-family:var(--xve-font, var(--office-font, system-ui, sans-serif));
	-webkit-user-select:none;user-select:none;contain:strict}
.xg-scroller{position:absolute;inset:0;overflow:auto;overscroll-behavior:contain;touch-action:pan-x pan-y}
.xg-space{position:relative}
.xg-view{position:sticky;left:0;top:0;overflow:hidden;cursor:cell;touch-action:pan-x pan-y}
.xg-q,.xg-hbox{position:absolute;overflow:hidden}
.xg-plane{position:absolute;left:0;top:0;will-change:transform}
.xg-cells,.xg-lines,.xg-overlay,.xg-drawings{position:absolute;left:0;top:0}
.xg-gl{position:absolute;background:var(--xve-grid-line, var(--office-border, #e1e1e1))}
.xg-c{position:absolute;box-sizing:border-box;overflow:visible}
.xg-c[hidden],.xg-gl[hidden],.xg-hd[hidden]{display:none}
.xg-c-over{z-index:1}
.xg-t{position:absolute;top:0;bottom:0;display:flex;box-sizing:border-box;overflow:hidden;line-height:1.15}
.xg-t.xg-wrap{white-space:pre-wrap;overflow-wrap:anywhere}
.xg-t.xg-wrap>.xg-tx{display:block;width:100%}
.xg-t.xg-vert{writing-mode:vertical-rl;text-orientation:upright}
.xg-t.xg-over{background:var(--xve-sheet-bg, var(--office-paper, #fff))}
.xg-c[style*="background-color"]>.xg-t.xg-over,.xg-c-fill>.xg-t.xg-over{background:transparent}
.xg-tx{display:inline-block;white-space:inherit}
.xg-link{cursor:pointer}
.xg-b{position:absolute;left:-1px;top:-1px;box-sizing:border-box;pointer-events:none}
.xg-diag{position:absolute;left:0;top:0;pointer-events:none}
.xg-spark{position:absolute;left:0;top:0;pointer-events:none;overflow:hidden}
.xg-db{position:absolute;left:2px;top:2px;bottom:2px;border:1px solid;box-sizing:border-box}
.xg-ic{position:absolute;left:2px;top:50%;transform:translateY(-50%)}
.xg-cm{position:absolute;right:0;top:0;width:0;height:0;border-left:6px solid transparent;border-top:6px solid #d13438;z-index:2}
.xg-er{position:absolute;left:0;top:0;width:0;height:0;border-right:6px solid transparent;border-top:6px solid #107c41;z-index:2}
.xg-overlay{pointer-events:none;z-index:3}
.xg-shade{position:absolute;background:var(--xve-selection, color-mix(in srgb, var(--office-accent, #217346) 12%, transparent))}
.xg-outline{position:absolute;box-sizing:border-box;border:2px solid var(--xve-selection-border, var(--office-accent, #217346))}
.xg-active{position:absolute;box-sizing:border-box;border:1px solid var(--xve-selection-border, var(--office-accent, #217346))}
.xg-handle{position:absolute;box-sizing:border-box;background:var(--xve-selection-border, var(--office-accent, #217346));border:1px solid #fff;pointer-events:auto;cursor:crosshair}
.xg-ants{position:absolute;box-sizing:border-box;--a:var(--xve-selection-border, var(--office-accent, #217346));
	background-image:linear-gradient(90deg,var(--a) 50%,transparent 50%),linear-gradient(90deg,var(--a) 50%,transparent 50%),linear-gradient(0deg,var(--a) 50%,transparent 50%),linear-gradient(0deg,var(--a) 50%,transparent 50%);
	background-repeat:repeat-x,repeat-x,repeat-y,repeat-y;background-size:8px 2px,8px 2px,2px 8px,2px 8px;
	background-position:0 0,0 100%,0 0,100% 0;animation:xg-ants .5s linear infinite}
@keyframes xg-ants{to{background-position:8px 0,-8px 100%,0 -8px,100% 8px}}
@media (prefers-reduced-motion:reduce){.xg-ants{animation:none}}
.xg-fill-preview{position:absolute;box-sizing:border-box;border:1px dashed #666}
.xg-refbox{position:absolute;box-sizing:border-box;border:2px solid}
.xg-remote-layer{position:absolute;left:0;top:0;pointer-events:none;z-index:3}
.xg-remote{position:absolute;box-sizing:border-box;border:2px solid}
.xg-remote[hidden]{display:none}
.xg-remote-tag{position:absolute;left:-2px;bottom:100%;max-width:160px;padding:0 4px;height:16px;line-height:16px;
	font:600 11px/16px var(--xve-font, var(--office-font, system-ui,sans-serif));color:#fff;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;border-radius:2px 2px 2px 0}
.xg-remote[data-tag="inside"] .xg-remote-tag{bottom:auto;top:0;left:0;border-radius:0 0 2px 0}
.xg-hbox{background:var(--xve-header-bg, var(--office-surface, #f3f3f3));z-index:4}
.xg-hdr-plane{position:absolute;left:0;top:0;will-change:transform}
.xg-hd{position:absolute;box-sizing:border-box;display:flex;align-items:center;justify-content:center;
	color:var(--xve-header-fg, var(--office-foreground, #444));border-right:1px solid var(--xve-header-line, var(--office-border, #d4d4d4));border-bottom:1px solid var(--xve-header-line, var(--office-border, #d4d4d4));overflow:hidden;white-space:nowrap}
.xg-hdr-col .xg-hd{top:0;height:100%}
.xg-hdr-row .xg-hd{left:0;width:100%}
.xg-hd-sel{background:var(--xve-header-active-bg, var(--office-selected, #e1e1e1));color:var(--xve-selection-border, var(--office-accent, #217346));font-weight:600}
.xg-hdr-col .xg-hd-sel{box-shadow:inset 0 -2px 0 var(--xve-selection-border, var(--office-accent, #217346))}
.xg-hdr-row .xg-hd-sel{box-shadow:inset -2px 0 0 var(--xve-selection-border, var(--office-accent, #217346))}
.xg-hd-full{background:var(--xve-header-full-bg, var(--office-selected, #cde6d7));color:var(--xve-selection-border, var(--office-accent, #185c37));font-weight:600}
.xg-hdr-col{cursor:s-resize}.xg-hdr-row{cursor:e-resize}
.xg-hdr-col .xg-hd,.xg-hdr-row .xg-hd{cursor:inherit}
.xg-corner{position:absolute;left:0;top:0;background:var(--xve-header-bg, var(--office-surface, #f3f3f3));z-index:5;box-sizing:border-box;
	border-right:1px solid var(--xve-header-line, var(--office-border, #d4d4d4));border-bottom:1px solid var(--xve-header-line, var(--office-border, #d4d4d4));cursor:default}
.xg-corner::after{content:"";position:absolute;right:3px;bottom:3px;border-left:9px solid transparent;border-bottom:9px solid #b4b4b4}
.xg-freeze-h,.xg-freeze-v{position:absolute;background:var(--xve-freeze-line, var(--office-border, #9b9b9b));z-index:6;pointer-events:none}
.xg-freeze-h.xg-split-bar,.xg-freeze-v.xg-split-bar{pointer-events:auto;background:var(--xve-split-bar, var(--office-border, #b4b4b4));touch-action:none}
.xg-freeze-h.xg-split-bar{cursor:row-resize}.xg-freeze-v.xg-split-bar{cursor:col-resize}
.xg-col-resize{cursor:col-resize!important}.xg-row-resize{cursor:row-resize!important}
.xg-resize-guide{position:absolute;z-index:8;border:0 dashed #666;pointer-events:none}
.xg-editor-host{position:absolute;z-index:7;box-sizing:border-box;opacity:0;overflow:hidden;pointer-events:none;background:#fff;color:#000;line-height:1.25}
.xg-editor-host.xg-editing{opacity:1;pointer-events:auto;overflow:visible;cursor:text;
	box-shadow:0 0 0 2px var(--xve-selection-border, var(--office-accent, #217346)),0 2px 6px rgba(0,0,0,.25)}
.xg-editor-host .xg-field-backdrop,.xg-editor-host .xg-field-input{padding:1px 3px}
.xg-sink{position:absolute;left:0;top:0;width:1px;height:1px;opacity:0;padding:0;border:0;resize:none;overflow:hidden;pointer-events:none}
.xg-field{position:relative}
.xg-field-backdrop,.xg-field-input{margin:0;border:0;box-sizing:border-box;font:inherit;line-height:inherit;letter-spacing:inherit;
	padding:inherit;white-space:pre-wrap;overflow-wrap:anywhere;word-break:normal;tab-size:4}
.xg-field-backdrop{position:absolute;inset:0;overflow:hidden;color:inherit;pointer-events:none;visibility:hidden}
.xg-field.xg-formula .xg-field-backdrop{visibility:visible}
.xg-field-input{position:relative;display:block;width:100%;height:100%;resize:none;outline:none;background:transparent;color:inherit;overflow:hidden}
.xg-field.xg-formula .xg-field-input{color:transparent;caret-color:#000}
.xg-ref{font:inherit}
.xg-editor-host .xg-field{width:100%;height:100%}
.xg-dd{position:absolute;z-index:6;width:17px;height:17px;box-sizing:border-box;border:1px solid #c8c8c8;background:#f3f3f3;
	display:flex;align-items:center;justify-content:center;cursor:default;pointer-events:auto}
.xg-dd::before{content:"";border-left:4px solid transparent;border-right:4px solid transparent;border-top:5px solid #444}
.xg-fb{position:absolute;z-index:6;box-sizing:border-box;padding:0;border:1px solid #c8c8c8;background:#f3f3f3;display:flex;align-items:center;justify-content:center;cursor:default;pointer-events:auto}
.xg-fb::before{content:"";border-left:4px solid transparent;border-right:4px solid transparent;border-top:5px solid #444}
.xg-fb-active{background:#dcebe1;border-color:var(--xve-selection-border, var(--office-accent, #217346))}
.xg-fb-active::before{border:0;width:8px;height:8px;background:var(--xve-selection-border, var(--office-accent, #217346));clip-path:polygon(0 0,100% 0,62% 50%,62% 100%,38% 85%,38% 50%)}
.xg-fmenu{padding:4px 0;font-size:13px;max-height:calc(100% - 8px)}
.xg-fm-row{position:relative}
.xg-fm-sep{border-top:1px solid var(--xve-border, var(--office-border, #d8d8d8));margin-top:3px;padding-top:3px}
.xg-fm-item{display:block;width:100%;text-align:left;padding:4px 14px;border:0;background:transparent;color:inherit;font:inherit;cursor:default}
.xg-fm-item:hover:not(:disabled),.xg-fm-item:focus-visible{background:var(--xve-accent, var(--office-selected, #e1efe6));outline:none}
.xg-fm-item:disabled{opacity:.45}
.xg-fm-parent::after{content:"";position:absolute;right:10px;top:50%;margin-top:-4px;border-top:4px solid transparent;border-bottom:4px solid transparent;border-left:5px solid currentColor}
.xg-fm-sub{position:absolute;left:100%;top:-4px;min-width:200px;padding:4px 0;max-height:360px;overflow:auto}
.xg-fm-sub[hidden]{display:none}
.xg-fm-head{padding:4px 14px;font-weight:600}
.xg-fm-swatches{display:flex;flex-wrap:wrap;gap:3px;padding:2px 14px 6px;max-width:220px}
.xg-fm-swatch{width:18px;height:18px;border:1px solid #888;padding:0}
.xg-fm-values{padding:4px 10px 6px;border-top:1px solid var(--xve-border, var(--office-border, #d8d8d8));margin-top:3px}
.xg-fm-search{box-sizing:border-box;width:100%;margin:4px 0;padding:3px 6px;font:inherit;border:1px solid var(--xve-border, var(--office-border, #c8c8c8));background:transparent;color:inherit}
.xg-fm-list{list-style:none;margin:2px 0;padding:0 0 0 4px;max-height:180px;overflow:auto;border:1px solid var(--xve-border, var(--office-border, #d8d8d8))}
.xg-fm-check{display:flex;align-items:center;gap:6px;padding:2px 4px;white-space:nowrap}
.xg-fm-empty{padding:4px;opacity:.7}
.xg-fm-buttons{display:flex;gap:6px;justify-content:flex-end;margin-top:6px}
.xg-fm-buttons button{min-width:64px;padding:3px 10px;border:1px solid var(--xve-border, var(--office-border, #c8c8c8));border-radius:3px;background:var(--xve-secondary, var(--office-surface, #f3f3f3));color:inherit;font:inherit}
.xg-fm-buttons .xg-fm-ok{background:var(--xve-selection-border, var(--office-accent, #217346));color:#fff;border-color:transparent}
.xg-popup{position:absolute;z-index:20;background:var(--xve-popover, var(--office-popover, #fff));color:var(--xve-popover-foreground, var(--office-popover-foreground, #222));
	border:1px solid var(--xve-border, var(--office-border, #c8c8c8));box-shadow:0 4px 12px rgba(0,0,0,.18);font-size:13px;border-radius:4px}
.xg-list{max-height:220px;overflow:auto;min-width:120px;padding:2px 0;margin:0;list-style:none}
.xg-list [role=option]{padding:3px 10px;cursor:default;white-space:nowrap}
.xg-list [aria-selected=true]{background:var(--xve-accent, var(--office-selected, #e1efe6))}
.xg-tip{padding:6px 8px;max-width:260px;white-space:pre-wrap;background:#fffbe6;color:#222;pointer-events:none}
.xg-tip b{display:block;margin-bottom:2px}
.xg-alert{padding:12px 14px;max-width:340px}
.xg-alert h3{margin:0 0 6px;font-size:14px}
.xg-alert p{margin:0 0 10px;white-space:pre-wrap}
.xg-alert .xg-actions{display:flex;gap:6px;justify-content:flex-end}
.xg-alert button{min-width:64px;padding:4px 10px;border:1px solid var(--xve-border, var(--office-border, #c8c8c8));border-radius:3px;background:var(--xve-secondary, var(--office-surface, #f3f3f3));color:inherit;font:inherit}
.xg-alert button:first-child{background:var(--xve-selection-border, var(--office-accent, #217346));color:#fff;border-color:transparent}
.xg-drawings{z-index:4}
.xg-obj{position:absolute;box-sizing:border-box;pointer-events:auto;cursor:default;touch-action:none}
.xg-obj img,.xg-obj svg{display:block;width:100%;height:100%;pointer-events:none}
.xg-chart svg [data-chart-series],.xg-chart svg [data-chart-part]{pointer-events:auto}
.xg-obj.xg-chart{background:#fff;outline:1px solid #d9d9d9;outline-offset:-1px}
.xg-obj.xg-obj-sel{outline:1px solid #8a8a8a;outline-offset:0}
.xg-obj .xg-grip{position:absolute;width:8px;height:8px;background:#fff;border:1px solid #555;box-sizing:border-box;border-radius:50%}
@media(pointer:coarse){.xg-grip::before{content:'';position:absolute;inset:-${(MIN_TOUCH_TARGET_PX - 8) / 2}px;touch-action:none}}
.xg-obj.xg-shape svg{overflow:visible}
.xg-obj.xg-formcontrol{overflow:visible}
.xg-formcontrol svg [data-control-part]{pointer-events:auto;cursor:pointer}
.xg-formcontrol[data-control-type=button] svg{pointer-events:auto}
.xg-fc-pop{position:absolute;left:0;z-index:5;max-height:160px;overflow:auto;background:#fff;border:1px solid #7a7a7a;box-sizing:border-box;font:12px 'Segoe UI',Arial,sans-serif}
.xg-fc-item{padding:2px 4px;cursor:pointer;white-space:nowrap}
.xg-fc-item:hover,.xg-fc-on{background:#0078d7;color:#fff}
.xg-obj-unsupported{background:repeating-linear-gradient(45deg,#f6f6f6,#f6f6f6 6px,#ececec 6px,#ececec 12px);border:1px dashed #aaa;
	font-size:11px;color:#555;display:flex;align-items:center;justify-content:center;text-align:center}
.xg-obj.xg-smartart{overflow:visible}
.xg-obj.xg-smartart office-ui-smartart{display:block;pointer-events:none}
.xg-smartart-placeholder{background:#fafafa;border:1px dashed #aaa;font-size:11px;color:#333;padding:4px;overflow:auto}
.xg-smartart-nodes{margin:0;padding-left:16px}
.xg-smartart-nodes .xg-smartart-child{margin-left:12px}
.xg-smartart-notice{display:none;position:absolute;left:0;bottom:0;padding:1px 4px;font-size:10px;background:rgba(255,255,255,.9);color:#555;border-top-right-radius:3px}
.xg-obj-sel .xg-smartart-notice,.xg-smartart-placeholder .xg-smartart-notice{display:block}
.xg-obj.xg-slicer{background:var(--sl-bg,#fff);border:1px solid var(--sl-border,#a6a6a6);overflow:hidden;
	font:12px/1.2 var(--xve-font, var(--office-font, "Segoe UI",system-ui,sans-serif));color:var(--sl-un-fg,#000)}
.xg-slicer button{font:inherit;color:inherit;margin:0;padding:0;box-sizing:border-box;cursor:pointer;touch-action:manipulation}
.xg-sl-head{display:flex;align-items:center;justify-content:space-between;box-sizing:border-box;padding:0 4px 0 8px;
	background:var(--sl-head-bg,#fff);color:var(--sl-head-fg,#000);border-bottom:1px solid var(--sl-head-line,#5b9bd5)}
.xg-sl-caption{font-weight:600;overflow:hidden;white-space:nowrap;text-overflow:ellipsis}
.xg-sl-tools{display:flex;gap:2px;flex:none}
.xg-sl-tool{width:20px;height:20px;display:flex;align-items:center;justify-content:center;background:transparent;border:1px solid transparent;border-radius:2px}
.xg-sl-tool:hover:not([disabled]){border-color:var(--sl-sel-line,#5b9bd5);background:var(--sl-sel-bg,#deebf7)}
.xg-sl-tool.xg-sl-on{background:var(--sl-sel-bg,#deebf7);border-color:var(--sl-sel-line,#5b9bd5)}
.xg-sl-tool[disabled]{opacity:.4;cursor:default}
@media(pointer:coarse){.xg-sl-tool{position:relative}.xg-sl-tool::before{content:'';position:absolute;inset:-${(MIN_TOUCH_TARGET_PX - 20) / 2}px}}
.xg-sl-body{position:absolute;overflow-x:hidden;overflow-y:auto}
.xg-sl-grid{position:relative;width:100%}
.xg-sl-item{position:absolute;display:block;text-align:left;padding:0 8px;overflow:hidden;white-space:nowrap;text-overflow:ellipsis;
	background:var(--sl-un-bg,#fff);color:var(--sl-un-fg,#000);border:1px solid var(--sl-un-line,#bfbfbf)}
.xg-sl-item:hover{background:var(--sl-hover,#eaf3fa);border-color:var(--sl-sel-line,#5b9bd5)}
.xg-sl-selected:hover,.xg-sl-selectedNoData:hover{background:var(--sl-sel-bg,#deebf7)}
.xg-sl-selected,.xg-sl-selectedNoData{background:var(--sl-sel-bg,#deebf7);color:var(--sl-sel-fg,#000);border-color:var(--sl-sel-line,#5b9bd5)}
.xg-sl-unselectedNoData,.xg-sl-selectedNoData{color:var(--sl-nd-fg,#8a8a8a)}
.xg-sl-selectedNoData{background:var(--sl-nd-bg,#f2f7fb)}
.xg-slicer button:focus-visible{outline:2px solid var(--xve-selection-border, var(--office-accent, #217346));outline-offset:-2px}
.xg-obj.xg-timeline{background:var(--sl-bg,#fff);border:1px solid var(--sl-border,#a6a6a6);overflow:hidden;
	font:12px/1.2 var(--xve-font, var(--office-font, "Segoe UI",system-ui,sans-serif));color:var(--sl-un-fg,#000)}
.xg-timeline button,.xg-timeline select{font:inherit;color:inherit;margin:0;box-sizing:border-box;cursor:pointer}
.xg-tl-head{display:flex;align-items:center;gap:8px;padding:0 4px 0 8px;border-bottom:1px solid var(--sl-head-line,#5b9bd5);box-sizing:border-box}
.xg-tl-caption{font-weight:600;overflow:hidden;white-space:nowrap;text-overflow:ellipsis}
.xg-tl-selection{flex:1;min-width:0;overflow:hidden;white-space:nowrap;text-overflow:ellipsis;color:var(--sl-nd-fg,#5a5a5a);text-align:center}
.xg-tl-level{height:22px;padding:0 4px;border:1px solid var(--sl-un-line,#bfbfbf);background:var(--sl-un-bg,#fff);border-radius:2px}
.xg-tl-body{position:absolute;overflow-x:auto;overflow-y:hidden}
.xg-tl-track{position:relative;height:100%;min-height:50px}
.xg-tl-group{position:absolute;top:0;box-sizing:border-box;padding:2px 0;border-left:1px solid var(--sl-un-line,#bfbfbf);color:var(--sl-un-fg,#000);font-weight:600;white-space:nowrap}
.xg-tl-group>span{position:sticky;left:0;padding:0 6px}
.xg-tl-cell:not(.xg-tl-on):hover{background:var(--sl-hover,#eaf3fa)}
.xg-tl-cell{position:absolute;padding:0;text-align:center;background:var(--sl-un-bg,#fff);border:1px solid var(--sl-un-line,#bfbfbf)}
.xg-tl-cell.xg-tl-on{background:var(--sl-sel-bg,#deebf7);border-color:var(--sl-sel-line,#5b9bd5)}
.xg-timeline button:focus-visible,.xg-timeline select:focus-visible{outline:2px solid var(--xve-selection-border, var(--office-accent, #217346));outline-offset:-2px}
.xg-live{position:absolute;width:1px;height:1px;overflow:hidden;clip:rect(0 0 0 0);white-space:nowrap}
`;
