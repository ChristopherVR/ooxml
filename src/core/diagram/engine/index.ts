// The per-point DiagramML layout engine: interprets a layout definition's `dgm:layoutNode` tree
// over the data model and places every presentation point (see `engine.ts`).
export { runSmartArtEngine, type EngineRun } from './engine';
export { runEngineLayout } from './engine-to-result';
export { ENGINE_FIRST_LAYOUT_IDS } from './engine-first-allowlist';
export { parseLayoutDefinitionXml } from './layout-def-parse';
export type { EngineNode } from './engine-node';
