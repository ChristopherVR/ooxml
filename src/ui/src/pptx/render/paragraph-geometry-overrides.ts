// Named, not `export *`: a star re-export of an external module inside the bundle becomes a
// runtime namespace copy that keeps every `ooxml-ui/pptx` export reachable for a host bundler.
export { resolveParagraphGeometryOverrides } from 'ooxml-core/pptx/editor/render/paragraph-geometry-overrides';
export type { ParagraphGeometryOverrides } from 'ooxml-core/pptx/editor/render/paragraph-geometry-overrides';
