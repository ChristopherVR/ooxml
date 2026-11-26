const BODY = '[data-pptx-editor-chrome] [data-pptx-chrome="body"]';
const COMMENTS = '[data-pptx-comments-panel]';
const PANE = `:is(${COMMENTS}, :has(${COMMENTS}))`;
const INSPECTOR = ':is([data-pptx-inspector], [data-pptx-chrome="inspector"])';

/** Keep the task pane in the right rail without taking a second column from the canvas. */
export const EDITOR_TASK_PANE_CSS = `
@media (min-width: 768px) {
  ${BODY} > ${PANE} {
    order: 1;
    flex: 0 0 clamp(224px, 30%, var(--office-side-pane-width, 320px));
    min-inline-size: 0 !important;
    min-block-size: 0;
  }
  ${BODY} > ${PANE} ${INSPECTOR} { inline-size: 100% !important; min-inline-size: 0; }
  ${BODY}:has(> ${PANE}) > :is(${INSPECTOR}, :has(> ${INSPECTOR})):not(${PANE}),
  ${BODY}:has(> ${PANE}) > [role="separator"]:has(+ :is(${PANE}, ${INSPECTOR})) {
    display: none;
  }
}
`;
