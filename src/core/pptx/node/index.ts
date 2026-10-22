/**
 * Node-only subpath export of the pptx area (`ooxml-core/pptx/node`).
 *
 * Holds the pieces of the engine that load Node-only packages, so the
 * browser-safe `ooxml-core/pptx` never does. Today that is the SVG
 * rasteriser of the `.ppt` writer:
 *
 * ```ts
 * import { PptxHandler } from 'ooxml-core/pptx';
 * import { rasterizeSvgInNode } from 'ooxml-core/pptx/node';
 *
 * const bytes = await handler.save(slides, { outputFormat: 'ppt', rasterizeSvg: rasterizeSvgInNode });
 * ```
 *
 * @module pptx/node
 */

export type { PptxSvgRasterizer } from '../core/core/types';
export { rasterizeSvgInNode } from './svg-rasterize-node';
