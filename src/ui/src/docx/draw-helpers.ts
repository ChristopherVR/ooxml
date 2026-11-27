// The editor's own helpers the Draw tab's code uses, handed over by `draw-loader.ts` when the tab
// loads. The editor bundles them already; importing them from the tab's lazy chunks as well would
// split them out of the editor's main chunk into chunks of their own, which every editor would then
// download separately, drawing or not.
import type { DocxInk } from 'ooxml-core/docx';
import type { Node as ProseMirrorNode } from 'prosemirror-model';
import type { locateFloat, nextFloatHeight } from './print-float-locate';

export interface DrawHelpers {
	locateFloat: typeof locateFloat;
	nextFloatHeight: typeof nextFloatHeight;
}

let helpers: (DrawHelpers & { inkOf(node: ProseMirrorNode): DocxInk | undefined }) | undefined;

/** Parsed `ink` attributes by text: the tools read the same drawings on every pointer move. */
const parsed = new Map<string, DocxInk | undefined>();

function inkOf(node: ProseMirrorNode): DocxInk | undefined {
	const attr = node.attrs.ink;
	if (typeof attr !== 'string' || node.type.name !== 'image') return undefined;
	if (!parsed.has(attr)) {
		if (parsed.size > 63) parsed.clear();
		let ink: DocxInk | undefined;
		try {
			ink = JSON.parse(attr) as DocxInk;
		} catch {
			ink = undefined;
		}
		parsed.set(attr, ink);
	}
	return parsed.get(attr);
}

export function setDrawHelpers(value: DrawHelpers): void {
	helpers = { ...value, inkOf };
}

/** The helpers; the Draw tab's code runs only after `draw-loader.ts` has set them. */
export function drawHelpers() {
	if (!helpers) throw new Error('The Draw tab ran before it was loaded.');
	return helpers;
}
