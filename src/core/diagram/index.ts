// Format-neutral SmartArt (DiagramML): the part model, DOM parsers for the data, layout, colours,
// quick-style and cached-drawing parts, relationship resolution and a loader. `pptx` and `docx`
// adapt to it; nothing here knows about slides or documents.
export * from './types.js';
export * from './attributes.js';
export * from './layout-category.js';
export * from './relationships.js';
export * from './data-model.js';
export * from './definitions.js';
export * from './drawing-color.js';
export * from './drawing-fill.js';
export * from './drawing-text.js';
export * from './drawing-geometry.js';
export * from './drawing.js';
export * from './load.js';
export { attributeReader } from './dom.js';
