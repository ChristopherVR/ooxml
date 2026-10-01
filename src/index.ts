// The root entry groups the shared building blocks by namespace. The document formats are
// symmetrical: `docx` and `pptx` are each imported through their own subpaths
// (`@christophervr/ooxml-core/docx`, `@christophervr/ooxml-core/pptx`), so importing the root never
// pulls in a format and its dependencies.
export * as color from './color/index.js';
export * as geometry from './geometry/index.js';
export * as opc from './opc/index.js';
export * as units from './units/index.js';
export * as xml from './xml/index.js';
