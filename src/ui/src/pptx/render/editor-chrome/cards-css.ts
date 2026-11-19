/**
 * The cards look (`data-office-look="cards"` on the editor or an ancestor, see `ooxml-ui/theme`):
 * the ribbon, thumbnail rail, canvas, notes and side pane float as cards on a workspace
 * background. Shape, elevation and density only; every colour comes from the viewer's palette.
 * Appended last to the chrome sheet so these rules win over the default look at equal weight.
 */
export const EDITOR_CARDS_ROOT =
	':is([data-office-look="cards"] [data-pptx-editor-chrome], [data-pptx-editor-chrome][data-office-look="cards"])';

const C = EDITOR_CARDS_ROOT;
const LIGHT_SHADOW = '0 1px 3px rgb(60 64 67 / 20%), 0 2px 6px rgb(60 64 67 / 8%)';
const FIELD =
	':is(input:not([type="checkbox"]):not([type="radio"]):not([type="color"]):not([type="range"]):not([type="file"]), select, pptx-ui-select)';
const SURFACES =
	'[data-pptx-chrome="slides"], [data-pptx-chrome="inspector"], [data-pptx-chrome="notes"]';
const SURFACE = `:is(${SURFACES})`;

export const EDITOR_CARDS_CSS = `
@media (min-width: 768px) {
  ${C} {
    --_pptx-card: var(--pptx-background);
    --_pptx-card-edge: var(--office-card-border, transparent);
    --_pptx-card-shadow: var(--office-shadow-card, ${LIGHT_SHADOW});
    --_pptx-gap: var(--office-workspace-gap, 12px);
    /* Light: a step darker than the cards; dark: a step lighter than the background. */
    background: var(--office-workspace, color-mix(in oklab, var(--pptx-secondary) 55%, var(--pptx-background)));
  }
  ${C} [data-pptx-title-bar] { background: transparent; border-bottom-color: transparent; }
  ${C} [data-pptx-chrome="ribbon"] {
    margin: 0 var(--_pptx-gap); box-sizing: border-box;
    border: 1px solid var(--_pptx-card-edge); border-radius: var(--office-radius-toolbar, 12px);
    background: var(--_pptx-card); box-shadow: var(--_pptx-card-shadow);
  }
  ${C} [data-pptx-chrome="ribbon"] :is([data-pptx-chrome="ribbon-primary"], [data-pptx-chrome="ribbon-tabs"], [data-pptx-chrome="ribbon-content"]) {
    background: transparent;
  }
  ${C} [data-pptx-chrome="ribbon-primary"] { border-radius: var(--office-radius-toolbar, 12px) var(--office-radius-toolbar, 12px) 0 0; }
  ${C} [data-pptx-chrome="body"] {
    gap: var(--_pptx-gap); padding: var(--_pptx-gap) var(--_pptx-gap) 0; box-sizing: border-box;
  }
  /* A splitter covers the gap it sits in instead of adding to it. */
  ${C} [data-pptx-chrome="body"] > [role="separator"] { margin-inline: calc(-1 * var(--_pptx-gap)); }
  ${C} ${SURFACE} {
    box-sizing: border-box; border: 1px solid var(--_pptx-card-edge);
    border-radius: var(--office-radius-surface, 14px);
    background: var(--_pptx-card); box-shadow: var(--_pptx-card-shadow);
  }
  ${C} [data-pptx-chrome="slides"] { overflow: hidden; }
  ${C} [data-pptx-chrome="slides"] [data-pptx-chrome="slide-list"] { padding-top: 8px; }
  ${C} [data-pptx-chrome="slide-row"] { border-radius: var(--office-radius-control, 7px); }
  ${C} [data-pptx-chrome="slide-frame"] { border-radius: 4px; }
  ${C} [data-pptx-chrome="slide-footer"] { border-top-color: transparent; }
  ${C} [data-pptx-chrome="body"] :has(> [data-pptx-viewport]) { background: transparent; }
  ${C} [data-pptx-chrome="body"] [data-pptx-viewport] {
    box-sizing: border-box; border: 1px solid var(--_pptx-card-edge);
    border-radius: var(--office-radius-surface, 14px);
    background: var(--_pptx-card); box-shadow: var(--_pptx-card-shadow);
  }
  ${C} [data-pptx-chrome="notes"] { width: auto; margin: var(--_pptx-gap) var(--_pptx-gap) 0; overflow: hidden; }
  ${C} [data-pptx-chrome="notes-header"] { padding-inline: 14px; }
  ${C} pptx-ui-status-bar { margin-top: 4px; background: transparent; border-top-color: transparent; }
  ${C} pptx-ui-status-bar::part(bar) { padding-inline: calc(var(--_pptx-gap) + 2px); }

  /* The side pane: 320px, a segmented tab strip and a roomier Format pane. */
  ${C} [data-pptx-chrome="body"] > :has(> [data-pptx-chrome="inspector"]),
  ${C} [data-pptx-chrome="inspector"] { min-width: var(--office-side-pane-width, 320px); }
  ${C} [data-pptx-chrome="inspector-header"] {
    padding: 10px 12px; border-bottom-color: color-mix(in oklab, var(--pptx-border) 60%, transparent);
  }
  ${C} [data-pptx-chrome="inspector-tabs"] {
    padding: 3px; border-radius: var(--office-radius-control, 7px); background: var(--pptx-accent);
  }
  ${C} [data-pptx-chrome="inspector-tabs"] button { border-radius: 5px; min-height: 28px; }
  ${C} [data-pptx-chrome="inspector-tabs"] button[aria-selected="true"] {
    /* Important: one binding marks its selected tab with an important accent fill. */
    background: var(--_pptx-card) !important; color: var(--pptx-foreground) !important;
    box-shadow: var(--_pptx-card-shadow), inset 0 0 0 1px var(--_pptx-card-edge);
  }
  ${C} [data-pptx-chrome="inspector"] { font-variant-numeric: tabular-nums; }
  ${C} [data-pptx-chrome="inspector-body"] { padding: 4px 14px 14px; }
  ${C} [data-pptx-chrome="inspector"] :is(pptx-ui-format-section, office-ui-format-section) {
    border-bottom-color: color-mix(in oklab, var(--pptx-border) 70%, transparent);
  }
  ${C} :is([data-pptx-chrome="inspector-heading"], .pptx-editor-heading) {
    font-size: 11px; line-height: 16px; font-weight: 600; letter-spacing: 0.06em; text-transform: uppercase;
  }
  ${C} [data-pptx-chrome="inspector"] :is([data-pptx-chrome="inspector-card"], .pptx-editor-card) {
    padding: 6px 0 10px; border: 0; border-radius: 0; background: transparent;
  }
  ${C} [data-pptx-chrome="inspector"] ${FIELD} {
    box-sizing: border-box; height: var(--office-pane-field-height, 32px);
    min-height: var(--office-pane-field-height, 32px);
    border-radius: var(--office-radius-control, 7px);
  }
  ${C} [data-pptx-chrome="inspector"] :is(.grid, [data-pptx-chrome="presentation-fields"]) { gap: 8px 10px; }
  ${C} :is(pptx-ui-theme-editor, pptx-ui-chart-format-pane):not([inline]) {
    overflow: hidden; border-left-color: transparent;
    border-radius: var(--office-radius-surface, 14px);
    box-shadow: var(--office-shadow-card-strong, ${LIGHT_SHADOW});
  }
  ${C} [data-pptx-chrome="theme-editor"] {
    border-left: 0; border-radius: var(--office-radius-surface, 14px);
    box-shadow: var(--office-shadow-card-strong, ${LIGHT_SHADOW});
  }
}
@media (min-width: 768px) and (forced-colors: active) {
  ${C} :is(${SURFACES}, [data-pptx-chrome="ribbon"], [data-pptx-viewport]) { border-color: CanvasText; }
  ${C} [data-pptx-chrome="inspector-tabs"] button[aria-selected="true"] {
    forced-color-adjust: none; background: Highlight !important; color: HighlightText !important;
    box-shadow: none;
  }
}
`;
