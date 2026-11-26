/**
 * The InkML runtime Word and Excel load on first use (`#inkml-runtime`, built as its own entry so
 * documents without ink never download it): the part reader and writer, and stroke edits.
 *
 * @module drawingml/inkml/runtime
 */
export { inkStrokesEqual, readInkPart, updatedInkPart, writeInkPart } from './part';
export { transformInkStrokes } from './edit';
