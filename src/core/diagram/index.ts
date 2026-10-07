// Format-neutral SmartArt (DiagramML): the part model, DOM parsers for the data, layout, colours,
// quick-style and cached-drawing parts, relationship resolution and a loader. `pptx` and `docx`
// adapt to it; nothing here knows about slides or documents.
export * from './types';
export * from './attributes';
export * from './layout-category';
export * from './relationships';
export * from './data-model';
export * from './definitions';
export * from './drawing-color';
export * from './drawing-fill';
export * from './drawing-text';
export * from './drawing-geometry';
export * from './drawing';
export * from './drawing-bounds';
export * from './load';
export { attributeReader } from './dom';
export * from './write-color';
export * from './write-fill';
