/**
 * Host-replaced regions (`customization.regions`): the container fills the
 * built-in's slot, and the built-in's own content is hidden beside it.
 */
export const REGION_CSS = `
.pptxv-region { display: block; box-sizing: border-box; width: 100%; height: 100%; min-width: 0; min-height: 0; overflow: auto; }
.pptxv-region-replaced > :not(.pptxv-region) { display: none !important; }
.pptxv-notes.pptxv-region-replaced[data-collapsed='true'] { display: none; }
.pptxv-statusbar.pptxv-region { height: auto; flex: none; }
.pptxv-modal-panel > .pptxv-region,
.pptxv-parity-dialog > .pptxv-region { height: auto; flex: 1 1 auto; }
.pptxv-backstage main > .pptxv-region { height: auto; flex: 1 1 auto; }
`;
