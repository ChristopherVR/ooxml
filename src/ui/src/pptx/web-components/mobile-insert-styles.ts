import { TABLE_PICKER_CSS } from './ribbon-insert-table-picker';

export const MOBILE_INSERT_STYLES = `
:host { display: block; min-width: 0; color: var(--pptx-foreground, #e2e8f0); font: 13px/1.3 system-ui, sans-serif; }
:host([hidden]), [hidden] { display: none !important; }
* { box-sizing: border-box; }
svg { flex: none; width: 26px; height: 26px; fill: none; stroke: currentColor; stroke-width: 1.5; stroke-linecap: round; stroke-linejoin: round; }
button { font: inherit; color: inherit; touch-action: manipulation; cursor: pointer; }
button:disabled { opacity: .4; cursor: default; }
button:focus-visible { outline: 2px solid var(--pptx-ring, #6366f1); outline-offset: 2px; }
.grid { display: grid; grid-template-columns: repeat(3, minmax(0, 1fr)); gap: 8px; }
.tile { display: flex; flex-direction: column; align-items: center; justify-content: center; gap: 6px; min-height: 80px; padding: 10px 6px; border: 1px solid var(--pptx-border, #33334d); border-radius: 12px; background: var(--pptx-muted, #26263a); text-align: center; overflow-wrap: anywhere; }
.tile svg { color: var(--pptx-primary, #6366f1); }
.tile:active:not(:disabled) { background: var(--pptx-accent, #33334d); }
.tile .more { width: 10px; height: 10px; margin-top: -4px; opacity: .6; }
.head { display: flex; align-items: center; gap: 4px; margin-bottom: 8px; }
.back { display: inline-flex; align-items: center; gap: 2px; min-height: 44px; padding: 0 10px 0 4px; border: 0; border-radius: 8px; background: transparent; color: var(--pptx-primary, #6366f1); }
.back svg { width: 18px; height: 18px; }
.head strong { font-size: 14px; }
.shapes { display: grid; grid-template-columns: repeat(auto-fill, minmax(48px, 1fr)); gap: 6px; }
.shape { display: flex; align-items: center; justify-content: center; min-height: 48px; padding: 0; border: 1px solid var(--pptx-border, #33334d); border-radius: 8px; background: var(--pptx-muted, #26263a); }
.shape svg { width: 22px; height: 22px; }
.rows { display: flex; flex-direction: column; gap: 6px; }
.row { display: flex; align-items: center; gap: 10px; min-height: 48px; padding: 6px 12px; border: 1px solid var(--pptx-border, #33334d); border-radius: 10px; background: var(--pptx-muted, #26263a); text-align: left; }
.row[aria-current="true"] { border-color: var(--pptx-primary, #6366f1); color: var(--pptx-primary, #6366f1); }
.row svg { width: 22px; height: 22px; color: var(--pptx-primary, #6366f1); }
${TABLE_PICKER_CSS}
.table-picker { display: grid; justify-items: center; padding: 4px 0 8px; }
.table-picker-label { font-size: 13px; }
.table-picker-cell { width: min(30px, calc((100vw - 72px) / 10)); height: min(30px, calc((100vw - 72px) / 10)); }
@media (forced-colors: active) {
	.tile, .shape, .row { border-color: ButtonText; background: Canvas; }
	.tile svg, .row svg, .back { color: ButtonText; }
}
`;
