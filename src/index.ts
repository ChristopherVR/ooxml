// Each area is also a subpath entry (`@christophervr/ooxml-core/xml`, `/opc`, ...). The root entry
// groups them by namespace so unrelated areas never compete for the same export names.
export * as color from './color/index.js';
export * as geometry from './geometry/index.js';
export * as opc from './opc/index.js';
export * as units from './units/index.js';
export * as xml from './xml/index.js';
