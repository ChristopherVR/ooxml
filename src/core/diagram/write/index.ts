// Format-neutral SmartArt part writers. Today the data part of a SmartArt graphic created in
// memory (`buildFabricatedDiagramDataXml`); the cached drawing and the definition parts are still
// written by pptx (`pptx/core/core/runtime/smartart-fabrication-*`).
export * from './data';
export * from './data-edit';
export * from './inserted';
