export const EDITOR_RIBBON_CSS = `
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
    min-height: 92px; padding: 2px 4px; gap: 0; box-sizing: border-box;
  }
  [data-pptx-editor-chrome] [data-pptx-chrome="ribbon-group-label"] {
    font-size: 9px; line-height: 9px; font-weight: 400;
    color: var(--pptx-muted-foreground); text-align: center;
  }
}
`;
