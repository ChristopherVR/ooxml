// Formula bar CSS (Excel for the web look), injected once into the shadow root.
export const FORMULA_BAR_CSS = `
.xfb{position:relative;display:flex;align-items:stretch;gap:0;min-height:26px;padding:3px 6px;box-sizing:border-box;background:var(--xve-background,#fff);border-bottom:1px solid var(--xve-border,#e1dfdd);font:12px/1.4 'Segoe UI',system-ui,sans-serif;color:var(--xve-foreground,#252423)}
.xfb-namebox{position:relative;display:flex;align-items:center;width:110px;flex:none;border:1px solid var(--xve-input,#c8c6c4);border-radius:var(--xve-radius,2px);background:var(--xve-background,#fff)}
.xfb-name{flex:1;min-width:0;border:0;outline:0;background:transparent;color:inherit;font:inherit;padding:0 4px;height:20px}
.xfb-namebox:focus-within{border-color:var(--xve-ring,#217346)}
.xfb-name-toggle{border:0;background:transparent;color:inherit;width:16px;height:20px;padding:0;display:flex;align-items:center;justify-content:center;cursor:pointer}
.xfb-name-list{position:absolute;left:-1px;top:100%;z-index:50;min-width:100%;max-height:240px;overflow:auto;margin:2px 0 0;padding:2px 0;list-style:none;background:var(--xve-popover,#fff);border:1px solid var(--xve-border,#c8c6c4);box-shadow:0 4px 12px rgba(0,0,0,.16)}
.xfb-name-item{padding:3px 8px;cursor:default;white-space:nowrap}
.xfb-name-item:hover{background:var(--xve-accent,#e1dfdd)}
.xfb-name-empty{padding:3px 8px;color:var(--xve-muted-foreground,#605e5c)}
.xfb-actions{display:flex;align-items:center;flex:none;padding:0 4px;margin-left:4px;border-left:1px solid var(--xve-border,#e1dfdd)}
.xfb-btn{border:0;background:transparent;color:inherit;width:24px;height:22px;padding:0;display:flex;align-items:center;justify-content:center;border-radius:var(--xve-radius,2px);cursor:pointer}
.xfb-btn:hover:not(:disabled){background:var(--xve-accent,#edebe9)}
.xfb-btn:disabled{opacity:.45;cursor:default}
.xfb-cancel,.xfb-enter{display:none}
.xfb-editing .xfb-cancel,.xfb-editing .xfb-enter{display:flex}
.xfb-fx i{font-family:Cambria,'Times New Roman',serif;font-size:13px}
.xfb-input{flex:1;min-width:0;height:20px;border:1px solid transparent;border-radius:var(--xve-radius,2px)}
.xfb-input:focus-within{border-color:var(--xve-input,#c8c6c4)}
.xfb-expanded .xfb-input{height:76px}
.xfb-expand{align-self:flex-start;flex:none;margin-left:2px}
.xfb-expanded .xfb-expand svg{transform:rotate(180deg)}
.xfb .xg-field{position:relative;overflow:hidden}
.xfb .xg-field-backdrop,.xfb .xg-field-input{position:absolute;inset:0;box-sizing:border-box;margin:0;border:0;padding:1px 4px;font:13px/18px Calibri,Carlito,'Segoe UI',sans-serif;white-space:pre-wrap;word-wrap:break-word;overflow:hidden}
.xfb .xg-field-backdrop{color:var(--xve-foreground,#252423);pointer-events:none;visibility:hidden}
.xfb .xg-field-input{resize:none;outline:0;background:transparent;color:inherit}
.xfb-expanded .xg-field-input{overflow:auto}
.xfb .xg-formula .xg-field-backdrop{visibility:visible}
.xfb .xg-formula .xg-field-input{color:transparent;caret-color:var(--xve-foreground,#252423)}
.xfb .xg-assist{top:100%}
`;
