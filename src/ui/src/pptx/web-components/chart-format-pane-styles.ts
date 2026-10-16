/** Docked like the theme editor: right edge of the editor body, a bottom sheet below 768px. */
export const CHART_FORMAT_PANE_STYLES = `
:host { display: block; position: fixed; z-index: 1090; box-sizing: border-box; width: 288px;
 color: var(--pptx-foreground, #e2e8f0); background: var(--pptx-card, #181c20);
 border-left: 1px solid var(--pptx-border, #33334d); box-shadow: -8px 0 24px #0002;
 font: 12px/1.4 var(--pptx-ui-font-family, system-ui, sans-serif); }
:host([inline]) { position: static; width: 100%; border: 0; box-shadow: none; }
* { box-sizing: border-box; }
.panel { display: flex; flex-direction: column; height: 100%; min-height: 0; }
header { display: flex; align-items: center; justify-content: space-between; gap: 8px; padding: 10px 10px 6px; flex: none; }
h3 { margin: 0; font-size: 14px; font-weight: 600; }
.options-row { display: flex; align-items: end; gap: 6px; padding: 0 10px 6px; flex: none; }
.picker { flex: 1; min-width: 0; }
.picker > span, .caption { display: block; color: var(--pptx-muted-foreground, #94a3b8); font-size: 11px; }
.picker > span { margin-bottom: 3px; }
.tabs { display: flex; gap: 2px; padding: 0 10px; border-bottom: 1px solid var(--pptx-border, #33334d); flex: none; overflow-x: auto; }
.tabs button { border: 0; border-bottom: 2px solid transparent; border-radius: 0; background: transparent; padding: 6px 8px; white-space: nowrap; }
.tabs button[aria-selected=true] { border-bottom-color: var(--pptx-primary, #ec6b3d); font-weight: 600; }
.content { display: grid; align-content: start; gap: 4px; padding: 6px 10px 12px; overflow-y: auto; min-height: 0; flex: 1; }
section { border-bottom: 1px solid var(--pptx-border, #33334d); padding-bottom: 6px; }
.section-toggle { display: flex; width: 100%; align-items: center; gap: 6px; border: 0; background: transparent; padding: 6px 0; font-weight: 600; text-align: left; }
.section-toggle::before { content: ''; width: 0; height: 0; border-left: 5px solid currentColor; border-top: 4px solid transparent; border-bottom: 4px solid transparent; transition: transform .12s; }
.section-toggle[aria-expanded=true]::before { transform: rotate(90deg); }
.fields { display: grid; gap: 6px; padding: 2px 0 4px 11px; }
.radio-group { border: 0; margin: 0; padding: 0; display: grid; gap: 3px; }
.radio-group legend { position: absolute; width: 1px; height: 1px; overflow: hidden; clip: rect(0 0 0 0); }
.choice, .field-checkbox { display: flex; align-items: center; gap: 6px; }
.choice .caption, .field-checkbox .caption { color: inherit; font-size: 12px; }
.choice:has(pptx-ui-radio[disabled]) { opacity: .5; }
.field { display: grid; grid-template-columns: 1fr minmax(0, 112px) auto; align-items: center; gap: 6px; }
.field-checkbox { display: flex; }
.unit { color: var(--pptx-muted-foreground, #94a3b8); font-size: 11px; min-width: 14px; }
pptx-ui-select, input[type=number], input[type=text] { display: block; width: 100%; min-height: 26px; padding: 3px 6px; color: inherit; font: inherit;
 background: var(--pptx-background, #0e1114); border: 1px solid var(--pptx-border, #33334d); border-radius: 4px; }
input[type=color] { width: 44px; height: 26px; padding: 0; border: 1px solid var(--pptx-border, #33334d); background: transparent; justify-self: start; }
button { color: inherit; font: inherit; background: var(--pptx-secondary, #20252b);
 border: 1px solid var(--pptx-border, #33334d); border-radius: 4px; padding: 4px 8px; cursor: pointer; }
button:hover:not(:disabled) { background: var(--pptx-accent, #33334d); }
button:disabled, input:disabled { opacity: .5; cursor: default; }
.text-options[aria-pressed=true] { border-color: var(--pptx-primary, #ec6b3d); background: var(--pptx-accent, #33334d); }
.file-button { justify-self: start; }
.close { border: 0; background: transparent; min-width: 28px; min-height: 28px; font-size: 18px; line-height: 1; }
.note { margin: 2px 0 0; color: var(--pptx-muted-foreground, #94a3b8); font-size: 11px; }
button:focus-visible, input:focus-visible { outline: 2px solid var(--pptx-ring, #ec6b3d); outline-offset: 1px; }
@media (pointer: coarse), (max-width: 767px) {
 pptx-ui-select, input[type=number], input[type=text], .close, .tabs button { min-height: 44px; }
 pptx-ui-radio, pptx-ui-checkbox { width: 20px; height: 20px; }
 .choice, .field-checkbox { min-height: 32px; }
 :host(:not([inline])) { border-left: 0; border-top: 1px solid var(--pptx-border, #33334d); border-radius: 12px 12px 0 0; }
}
@media (forced-colors: active) { :host, section, button, input, pptx-ui-select { border-color: CanvasText; }
 .tabs button[aria-selected=true] { border-bottom-color: Highlight; } }
`;
