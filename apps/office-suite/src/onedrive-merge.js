/** Load only the current format's merger; the storage client stays independent of Office engines. */
export async function mergeOneDriveChanges({ kind, base, local, remote, signal }) {
	signal?.throwIfAborted();
	const merger =
		kind === 'docx'
			? (await import('ooxml-core/docx')).mergeDocxSnapshots
			: kind === 'xlsx'
				? (await import('ooxml-core/xlsx/collab')).mergeWorkbookSnapshots
				: kind === 'pptx'
					? (await import('ooxml-core/pptx')).mergePptxSnapshots
					: (await import('ooxml-core/visio/collab')).mergeDrawingSnapshots;
	signal?.throwIfAborted();
	const result = await merger(base.bytes, local, remote.bytes);
	signal?.throwIfAborted();
	return result;
}
