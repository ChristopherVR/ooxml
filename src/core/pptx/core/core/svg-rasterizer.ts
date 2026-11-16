// A leaf module (no imports), so strictly compiled code such as `pptx/node` can name the
// rasterizer without pulling in the relaxed pptx model; `types.ts` re-exports it.

/**
 * Rasterises SVG bytes to PNG for a `.ppt` save (the binary format has no SVG
 * picture). `sizeFor` maps the SVG's intrinsic size in pixels (0 when it
 * declares none) to the raster size the writer wants. Resolve `undefined`
 * when the SVG cannot be rasterised; the writer then writes a placeholder
 * with a warning. `ooxml-core/pptx/node` exports one for Node.js.
 */
export type PptxSvgRasterizer = (
	svg: Uint8Array,
	sizeFor: (intrinsicWidth: number, intrinsicHeight: number) => { width: number; height: number },
) => Promise<Uint8Array | undefined>;
