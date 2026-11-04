import { createStorageMerger } from 'ooxml-core/embedding';

const lazy =
	(load) =>
	async ({ base, local, remote, signal }) => {
		const merge = await load();
		signal?.throwIfAborted();
		return merge(base.bytes, local, remote.bytes);
	};

/** Hosts may replace individual handlers, disable them with undefined, or replace the router. */
export const officeMergeHandlers = {
	docx: lazy(async () => (await import('ooxml-core/docx')).mergeDocxSnapshots),
	xlsx: lazy(async () => (await import('ooxml-core/xlsx/collab')).mergeWorkbookSnapshots),
	pptx: lazy(async () => (await import('ooxml-core/pptx')).mergePptxSnapshots),
	vsdx: lazy(async () => (await import('ooxml-core/visio/collab')).mergeDrawingSnapshots),
};
export const mergeOneDriveChanges = createStorageMerger(officeMergeHandlers);
