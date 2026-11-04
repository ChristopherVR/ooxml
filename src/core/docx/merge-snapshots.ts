import { mergeOpcSnapshots, type SnapshotMergeResult } from '../opc/merge-snapshots';
import { mergeXmlThreeWay } from '../xml/merge-three-way';
import { NS } from '../xml/namespaces';

const WORD_2010 = 'http://schemas.microsoft.com/office/word/2010/wordml';
const TEXT_PART =
	/^word\/(?:document|header\d*|footer\d*|footnotes|endnotes|comments|styles)\.xml$/u;
const POSITIONAL = new Set(['body', 'p', 'r', 'hyperlink', 'sdtContent', 'tbl', 'tr', 'tc']);
const IDS: Readonly<Record<string, string>> = {
	comment: 'id',
	footnote: 'id',
	endnote: 'id',
	style: 'styleId',
};

function keyElement(element: Element): string | undefined {
	if (element.namespaceURI !== NS.w) return undefined;
	if (element.localName === 'p') {
		const id = element.getAttributeNS(WORD_2010, 'paraId');
		return id ? `paragraph:${id}` : undefined;
	}
	const attribute = IDS[element.localName];
	if (!attribute) return undefined;
	const id = element.getAttributeNS(NS.w, attribute);
	return id ? `${element.localName}:${id}` : undefined;
}

/**
 * Three-way merge of independent edits to native DOCX snapshots with a common base. This reads
 * source XML rather than reconstructing a document model, so unsupported one-sided parts survive.
 * Ambiguous structure, overlapping text edits and opaque same-part conflicts remain unresolved.
 * This does not synthesize the Yjs operation history absent from independently saved packages.
 */
export function mergeDocxSnapshots(
	base: Uint8Array,
	local: Uint8Array,
	remote: Uint8Array,
): Promise<SnapshotMergeResult> {
	return mergeOpcSnapshots(base, local, remote, {
		async mergePart(input) {
			if (!TEXT_PART.test(input.path) || !input.base || !input.local || !input.remote)
				return {
					status: 'unresolved',
					issues: [`Concurrent unsupported DOCX part: ${input.path}`],
				};
			const decoder = new TextDecoder();
			const result = mergeXmlThreeWay(
				decoder.decode(input.base),
				decoder.decode(input.local),
				decoder.decode(input.remote),
				{
					keyElement,
					allowPositional: (parent) =>
						parent.namespaceURI === NS.w && POSITIONAL.has(parent.localName),
					allowText: (element) =>
						(element.namespaceURI === NS.w || element.namespaceURI === NS.a) &&
						element.localName === 't',
				},
			);
			return result.status === 'merged'
				? { status: 'merged', bytes: new TextEncoder().encode(result.xml) }
				: { status: 'unresolved', issues: result.issues.map((issue) => `${input.path}: ${issue}`) };
		},
	});
}
