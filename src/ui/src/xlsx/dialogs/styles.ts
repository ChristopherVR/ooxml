// Styles for the command dialogs' light-DOM content (slotted into office-ui-dialog). Uses the
// editor's `--xve-*` tokens with neutral fallbacks so a dialog also reads well without the shell.
export const DIALOG_CSS = `
.xve-cmd-dialog { --office-background: var(--xlsx-theme-popover);
 --office-foreground: var(--xlsx-theme-popover-foreground);
 --office-border: var(--xlsx-theme-border); --office-ring: var(--xlsx-theme-ring); }
.xve-dialog-body { display: flex; flex-direction: column; gap: 10px; min-width: min(360px, 86vw);
	font: var(--office-font-size,13px)/var(--office-line-height,1.4) var(--xve-font, var(--office-font, system-ui, sans-serif)); color: var(--office-foreground); }
.xve-cmd-dialog-wide .xve-dialog-body { min-width: min(560px, 90vw); }
.xve-dialog-footer { display: flex; gap: 8px; justify-content: flex-end; flex-wrap: wrap; }
.xve-btn { min-height: 28px; padding: 3px 14px; border-radius: var(--xve-radius, var(--office-radius, 4px)); cursor: pointer;
	border: 1px solid var(--xve-border, var(--office-border, #c8c6c4)); background: var(--xve-secondary, var(--office-surface, #fff)); color: inherit; font: inherit; }
.xve-btn:hover:not(:disabled) { background: var(--xve-accent, var(--office-selected, #f3f2f1)); }
.xve-btn:focus-visible, .xve-input:focus-visible, .xve-tab:focus-visible, .xve-listbox:focus-visible,
	.xve-tile:focus-visible { outline: 2px solid var(--xve-ring, var(--office-ring, #217346)); outline-offset: 1px; }
.xve-btn-primary { background: var(--xve-primary, var(--office-accent, #217346)); color: var(--xve-primary-foreground, var(--office-accent-foreground, #fff));
	border-color: var(--xve-primary, var(--office-accent, #217346)); }
.xve-btn-primary:hover:not(:disabled) { background: var(--xve-primary, var(--office-accent, #1b5e38)); filter: brightness(.95); }
.xve-btn:disabled { opacity: .5; cursor: default; }
.xve-field { display: grid; grid-template-columns: minmax(110px, auto) 1fr; align-items: center; gap: 8px; }
.xve-field-label { color: var(--xve-muted-foreground, var(--office-muted-foreground, #605e5c)); }
.xve-input { box-sizing: border-box; min-height: 28px; padding: 3px 6px; border-radius: var(--xve-radius, var(--office-radius, 4px));
	border: 1px solid var(--xve-input, var(--office-border, #c8c6c4)); background: var(--xve-background, var(--office-background, #fff)); color: inherit; font: inherit; min-width: 0; }
textarea.xve-input { resize: vertical; }
.xve-input[aria-invalid="true"] { border-color: var(--xve-destructive, var(--office-danger, #c50f1f)); }
.xve-check { display: flex; align-items: center; gap: 6px; }
.xve-checklist { display: flex; flex-direction: column; gap: 4px; max-height: 240px; overflow: auto; padding: 4px 2px; }
.xve-fieldset { border: 1px solid var(--xve-border, var(--office-border, #e1dfdd)); border-radius: var(--xve-radius, var(--office-radius, 4px));
	padding: 8px 10px; margin: 0; display: flex; flex-direction: column; gap: 6px; min-width: 0; }
.xve-fieldset legend { padding: 0 4px; font-weight: 600; }
.xve-row { display: flex; gap: 10px; align-items: flex-start; flex-wrap: wrap; }
.xve-row > * { flex: 1 1 160px; }
.xve-note { margin: 0; color: var(--xve-muted-foreground, var(--office-muted-foreground, #605e5c)); font-size: 12px; }
.xve-tablist { display: flex; gap: 2px; border-bottom: 1px solid var(--xve-border, var(--office-border, #e1dfdd)); flex-wrap: wrap; }
.xve-tab { border: 0; background: transparent; color: inherit; font: inherit; padding: 6px 10px; cursor: pointer;
	border-bottom: 2px solid transparent; }
.xve-tab[aria-selected="true"] { border-bottom-color: var(--xve-primary, var(--office-accent, #217346)); font-weight: 600; }
.xve-tabpanel { padding-top: 10px; display: flex; flex-direction: column; gap: 10px; }
.xve-tabpanel[hidden] { display: none; }
.xve-listbox { border: 1px solid var(--xve-input, var(--office-border, #c8c6c4)); border-radius: var(--xve-radius, var(--office-radius, 4px));
	min-height: 120px; max-height: 220px; overflow: auto; background: var(--xve-background, var(--office-background, #fff)); }
.xve-option { padding: 3px 8px; cursor: default; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
.xve-option[aria-selected="true"] { background: var(--xve-selection, var(--office-selected, #cce8d6)); color: inherit; }
.xve-sample { min-height: 40px; display: flex; align-items: center; justify-content: center; padding: 6px;
	border: 1px solid var(--xve-border, var(--office-border, #e1dfdd)); background: var(--xve-sheet-bg, var(--office-paper, #fff)); color: #000; overflow: hidden; }
.xve-tiles { display: grid; grid-template-columns: repeat(auto-fill, minmax(84px, 1fr)); gap: 6px; }
.xve-tile { display: flex; flex-direction: column; align-items: center; gap: 4px; padding: 6px; cursor: pointer;
	border: 1px solid var(--xve-border, var(--office-border, #e1dfdd)); border-radius: var(--xve-radius, var(--office-radius, 4px)); background: transparent; color: inherit; font: inherit; }
.xve-tile[aria-pressed="true"] { border-color: var(--xve-primary, var(--office-accent, #217346)); background: var(--xve-selection, var(--office-selected, #e6f2ea)); }
.xve-symbols { display: grid; grid-template-columns: repeat(16, 28px); gap: 2px; max-height: 240px; overflow: auto; }
.xve-symbols button { width: 28px; height: 28px; padding: 0; font-size: 16px; }
.xve-preview { border: 1px solid var(--xve-border, var(--office-border, #e1dfdd)); background: #fff; min-height: 160px; display: flex;
	align-items: center; justify-content: center; }
.xve-preview svg { max-width: 100%; height: auto; }
.xve-results { max-height: 160px; overflow: auto; border: 1px solid var(--xve-border, var(--office-border, #e1dfdd)); }
.xve-results table { border-collapse: collapse; width: 100%; font-size: 12px; }
.xve-results th, .xve-results td { text-align: left; padding: 2px 6px; border-bottom: 1px solid var(--xve-border, var(--office-border, #f3f2f1)); }
.xve-results tr[aria-selected="true"] td { background: var(--xve-selection, var(--office-selected, #cce8d6)); }
.xve-border-box { display: grid; grid-template-columns: 1fr 1fr; gap: 10px; }
.xve-border-preview { width: 120px; height: 80px; position: relative; background: #fff; border: 1px dashed #c8c6c4; }
@media (max-width: 600px) { .xve-field { grid-template-columns: 1fr; } .xve-symbols { grid-template-columns: repeat(8, 28px); } }
`;
