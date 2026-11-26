const FORMAT_ARRANGE = `[data-pptx-editor-chrome] :is(
  [data-ribbon-group="videoFormat.arrange"],
  [data-ribbon-group="audioFormat.arrange"],
  [data-ribbon-group="smartArtFormat.arrange"]
)`;

export const EDITOR_RIBBON_CSS = `
[data-pptx-comments-panel] { position: relative; }
[data-pptx-comment-close] {
  position: absolute; inset: 4px 4px auto auto; z-index: 1; display: grid; place-items: center;
  width: 30px; height: 30px; padding: 0; border: 0; border-radius: 4px;
  color: inherit; background: var(--pptx-card); cursor: pointer;
}
[data-pptx-comment-close]:hover { background: var(--pptx-accent); }
[data-pptx-comment-close] svg { width: 16px; height: 16px; stroke: currentColor; stroke-width: 1.5; }
@media (max-width: 767px) {
  [data-pptx-comment-close] { width: 44px; height: 44px; }
}
[data-pptx-editor-chrome] .contextual-fields > pptx-ui-ribbon-gallery[mode="inline"] {
  grid-row: span 3;
}
[data-pptx-editor-chrome] .contextual-number {
  display: flex; align-items: center; justify-content: space-between;
  gap: 6px; font-size: 12px; white-space: nowrap;
}
[data-pptx-editor-chrome] .contextual-number > input {
  box-sizing: border-box; width: 68px; height: 24px; padding: 0 4px;
  border: 1px solid var(--pptx-border); border-radius: var(--pptx-radius, 4px);
  background: var(--pptx-background); color: var(--pptx-foreground);
}
[data-pptx-editor-chrome] .contextual-number > input:disabled { opacity: 0.5; }
[data-pptx-editor-chrome] .contextual-color {
  display: inline-flex; align-items: center; gap: 6px; min-height: 24px;
  padding: 0 6px; font-size: 12px; white-space: nowrap;
}
[data-pptx-editor-chrome] .contextual-color > input {
  box-sizing: border-box; width: 18px; height: 18px; padding: 0; border: 0; background: transparent;
}
@media (min-width: 901px) and (pointer: fine) {
  ${FORMAT_ARRANGE} :is(div:has(> [data-pptx-chrome="arrange-controls"]), [data-pptx-chrome="arrange-controls"], [data-pptx-chrome="arrange-extras"], [data-pptx-chrome="control-fragment"]) {
    display: contents;
  }
  ${FORMAT_ARRANGE} .pptx-rb-sep { display: none; }
}
@media (min-width: 768px) {
  [data-pptx-editor-chrome] [data-pptx-chrome="ribbon"] {
    background: color-mix(in oklab, var(--pptx-secondary) 50%, transparent);
    color: var(--pptx-foreground); border-bottom: 1px solid var(--pptx-border);
  }
  /* Full-screen mode lays the whole ribbon over the slide: the half-transparent fill above
     would let the slide show through, so the page background goes underneath it. */
  [data-pptx-editor-chrome] [data-pptx-chrome="ribbon"][data-ribbon-display="fullScreen"][data-ribbon-peek] {
    background:
      linear-gradient(
        color-mix(in oklab, var(--pptx-secondary) 50%, transparent),
        color-mix(in oklab, var(--pptx-secondary) 50%, transparent)
      ),
      var(--pptx-background);
  }
  [data-pptx-editor-chrome] [data-pptx-chrome="ribbon-primary"] {
    min-height: 32px; height: 32px; padding: 2px 6px; gap: 2px;
    border: 0; box-sizing: border-box;
  }
  /* The tab row is the shared \`pptx-ui-ribbon-tabs\`: its tabs and underline are drawn in its
     shadow root from --office-* tokens, so only the row's box is set here. */
  [data-pptx-editor-chrome] [data-pptx-chrome="ribbon-tabs"] {
    display: block; min-height: 35px; padding: 0 4px; box-sizing: border-box;
  }
  [data-pptx-editor-chrome] [data-pptx-chrome="ribbon-content"] {
    min-width: 0; max-width: 100%; min-height: 92px; padding: 2px 4px; gap: 0; box-sizing: border-box;
  }
  [data-pptx-editor-chrome] [data-pptx-chrome="ribbon-group-label"] {
    font-size: 9px; line-height: 9px; font-weight: 400;
    color: var(--pptx-muted-foreground); text-align: center;
  }
}
`;
