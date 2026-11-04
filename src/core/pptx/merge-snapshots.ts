import { mergeOpcSnapshots, type SnapshotMergeResult } from '../opc/merge-snapshots';
import { mergeXmlThreeWay } from '../xml/merge-three-way';
import { NS } from '../xml/namespaces';

const SLIDE_PART = /^ppt\/(?:slides\/slide\d+|notesSlides\/notesSlide\d+)\.xml$/u;
const SHAPES = new Set(['sp', 'pic', 'cxnSp', 'graphicFrame', 'grpSp']);
const POSITIONAL = new Set(['txBody', 'p', 'r', 'fld']);

function keyElement(element: Element, parent: Element): string | undefined {
	if (
		element.namespaceURI !== NS.p ||
		!SHAPES.has(element.localName) ||
		parent.namespaceURI !== NS.p ||
		(parent.localName !== 'spTree' && parent.localName !== 'grpSp')
	)
		return undefined;
	const id = element.getElementsByTagNameNS(NS.p, 'cNvPr')[0]?.getAttribute('id');
	return id ? `shape:${id}` : undefined;
}

/**
 * Merge independent edits to source-backed PPTX snapshots sharing a common base. Stable cNvPr
 * shape IDs align text and geometry/property edits without replacing unrelated slide content.
 * Repeated unkeyed text structures require unchanged alignment; ambiguous edits stay unresolved.
 * This is snapshot reconciliation, not a replacement for a live collaboration operation log.
 */
export function mergePptxSnapshots(
	base: Uint8Array,
	local: Uint8Array,
	remote: Uint8Array,
): Promise<SnapshotMergeResult> {
	return mergeOpcSnapshots(base, local, remote, {
		async mergePart(input) {
			if (!SLIDE_PART.test(input.path) || !input.base || !input.local || !input.remote)
				return {
					status: 'unresolved',
					issues: [`Concurrent unsupported PPTX part: ${input.path}`],
				};
			const decoder = new TextDecoder();
			const result = mergeXmlThreeWay(
				decoder.decode(input.base),
				decoder.decode(input.local),
				decoder.decode(input.remote),
				{
					keyElement,
					allowPositional: (parent) =>
						(parent.namespaceURI === NS.a || parent.namespaceURI === NS.p) &&
						POSITIONAL.has(parent.localName),
					allowText: (element) => element.namespaceURI === NS.a && element.localName === 't',
				},
			);
			return result.status === 'merged'
				? { status: 'merged', bytes: new TextEncoder().encode(result.xml) }
				: { status: 'unresolved', issues: result.issues.map((issue) => `${input.path}: ${issue}`) };
		},
	});
}
