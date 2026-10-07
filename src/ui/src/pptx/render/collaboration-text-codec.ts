// Compatibility exports: document operations live in ooxml-core.
export {
	isYTextLike,
	encodeTextBody,
	encodeSegmentsToDelta,
	decodeDelta,
	decodeTextBody,
} from 'ooxml-core/pptx/editor/render/collaboration-text-codec';
export type { DeltaOp, YTextLike } from 'ooxml-core/pptx/editor/render/collaboration-text-codec';
