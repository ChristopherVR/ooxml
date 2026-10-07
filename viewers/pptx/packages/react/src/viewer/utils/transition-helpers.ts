/**
 * Thin re-export shim: the direction/orientation resolvers and the
 * `RANDOM_ELIGIBLE_TYPES` / `INSTANT` constants now live in
 * `pptx-viewer-shared`. Kept so existing importers (`transition-resolver`,
 * tests) resolve unchanged.
 */
export type { ResolvedDirection, ResolvedDirection8 } from 'ooxml-ui/pptx';
export {
	resolveDirection,
	resolveDirection8,
	resolveOrientation,
	RANDOM_ELIGIBLE_TYPES,
	INSTANT,
} from 'ooxml-ui/pptx';
