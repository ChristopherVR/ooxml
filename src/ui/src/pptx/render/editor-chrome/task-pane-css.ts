const BODY = '[data-pptx-editor-chrome] [data-pptx-chrome="body"]';
const COMMENTS = '[data-pptx-comments-panel]';
const SELECTION = '[data-pptx-selection-pane]';
const PANE_CONTENT = `:is(${COMMENTS}, ${SELECTION})`;
const PANE = `:is(${PANE_CONTENT}, :has(${PANE_CONTENT}))`;
const INSPECTOR = ':is([data-pptx-inspector], [data-pptx-chrome="inspector"])';

/** Keep the task pane in the right rail without taking a second column from the canvas. */
export const EDITOR_TASK_PANE_CSS = `
${COMMENTS} { position: relative; }
${COMMENTS} > button[data-pptx-comment-close] {
  position: absolute; inset-block-start: 4px; inset-inline-end: 4px;
  z-index: 1; display: grid; place-items: center; box-sizing: border-box;
  inline-size: 30px; block-size: 30px; padding: 0; border: 0; border-radius: 4px;
  color: inherit; background: transparent; cursor: pointer;
}
${COMMENTS} > button[data-pptx-comment-close]:hover { background: var(--pptx-accent); }
${COMMENTS} > button[data-pptx-comment-close] svg {
  display: block; inline-size: 16px; block-size: 16px;
  fill: none; stroke: currentColor; stroke-width: 1.6; stroke-linecap: round;
}
${COMMENTS} button[data-pptx-comment-post] {
  align-self: flex-end; display: block; inline-size: fit-content; min-block-size: 28px;
  margin-inline-start: auto; padding: 4px 12px; border: 0; border-radius: 4px;
  background: var(--pptx-primary, var(--office-accent));
  color: var(--pptx-primary-foreground, var(--office-accent-foreground, white));
  font: inherit; font-weight: 600; cursor: pointer;
}
${COMMENTS} button[data-pptx-comment-post]:disabled { opacity: .45; cursor: default; }
${COMMENTS} button[data-pptx-comment-post]:hover:enabled { filter: brightness(.92); }
@media (max-width: 767px), (pointer: coarse) {
  ${COMMENTS} > button[data-pptx-comment-close] { inline-size: 44px; block-size: 44px; }
  ${COMMENTS} button[data-pptx-comment-post] { min-block-size: 44px; }
}
@media (min-width: 768px) {
  ${BODY} > ${PANE} {
    position: relative;
    order: 1;
    flex: 0 0 clamp(224px, 30%, var(--office-side-pane-width, 320px));
    min-inline-size: 0 !important;
    min-block-size: 0;
  }
  ${BODY} > ${PANE} :is(${INSPECTOR}, ${SELECTION}) { inline-size: 100% !important; min-inline-size: 0; }
  ${BODY}:has(${SELECTION}) > :is(${COMMENTS}, :has(${COMMENTS})):not(:has(${SELECTION})) {
    display: none;
  }
  ${BODY}:has(${PANE_CONTENT}) > :is(${INSPECTOR}, :has(> ${INSPECTOR})):not(${PANE}),
  ${BODY}:has(${PANE_CONTENT}) > [role="separator"]:has(+ ${INSPECTOR}) {
    display: none;
  }
}
`;
