import { PptxHandler } from '../index';
import type { PptxData } from '../index';

import type { ExecutionContext, TableStyleSaveOptions, ToolContext, ToolResult } from './types';

const PPTX_MIME_TYPE = 'application/vnd.openxmlformats-officedocument.presentationml.presentation';

/**
 * Load PptxData from a file, checking collaboration rooms first.
 * If a Y.Doc room is active the current in-memory state is dehydrated
 * so that unsaved edits are included in the tool's view of the document.
 */
export async function loadPresentation(
	filePath: string,
	execCtx: ExecutionContext,
): Promise<{ pptxData: PptxData; rawBytes: Uint8Array }> {
	const rawBytes = await execCtx.filesystem.readFile(filePath);

	// If collaboration is active, dehydrate from Y.Doc for current state
	if (execCtx.collaboration) {
		const room = execCtx.collaboration.getRoom(filePath);
		if (room) {
			const codec = execCtx.collaboration.getCodec(filePath);
			if (codec) {
				const currentBytes = await codec.dehydrate(room.ydoc);
				const handler = new PptxHandler();
				const pptxData = await handler.load(arrayBufferOf(currentBytes));
				return { pptxData, rawBytes };
			}
		}
	}

	// Fallback: parse from raw bytes on disk
	const handler = new PptxHandler();
	const pptxData = await handler.load(arrayBufferOf(rawBytes));
	return { pptxData, rawBytes };
}

/**
 * The bytes as a standalone ArrayBuffer (a view into a pooled buffer must not leak the rest). The copy
 * goes through the Uint8Array constructor because `slice()` on a Node Buffer is a view, not a copy.
 */
export function arrayBufferOf(bytes: Uint8Array): ArrayBuffer {
	return new Uint8Array(bytes).buffer;
}

/**
 * Saves a deck a tool changed, loading the original first so package parts the model does not
 * hold (relationships, custom XML, media) are kept.
 */
export async function saveDeckBytes(
	result: Pick<ToolResult, 'pptxData' | 'saveOptions'>,
	original: Uint8Array,
): Promise<Uint8Array> {
	const handler = new PptxHandler();
	await handler.load(arrayBufferOf(original));
	return handler.save(result.pptxData.slides, {
		headerFooter: result.pptxData.headerFooter,
		// `pptxData.tableStyleMap` carries any in-memory table-style edits a tool made (e.g.
		// `set_table_style_section`, `create_table_style`); without forwarding it here those edits
		// were silently dropped on save even though the core save pipeline supports them.
		tableStyles: result.pptxData.tableStyleMap,
		...result.saveOptions,
	});
}

/**
 * Save mutated PptxData back: routes through collaboration Y.Doc or disk.
 */
export async function savePresentation(
	filePath: string,
	pptxData: PptxData,
	rawBytes: Uint8Array,
	execCtx: ExecutionContext,
	tableStyleSaveOptions?: TableStyleSaveOptions,
): Promise<{ savedToDisk: boolean; routedThroughCollaboration: boolean }> {
	const outputBytes = await saveDeckBytes(
		{ pptxData, ...(tableStyleSaveOptions && { saveOptions: tableStyleSaveOptions }) },
		rawBytes,
	);

	// Collaboration mode: hydrate Y.Doc so peers receive the change
	if (execCtx.collaboration) {
		const room = execCtx.collaboration.getRoom(filePath);
		if (room) {
			const codec = execCtx.collaboration.getCodec(filePath);
			if (codec) {
				const origin = execCtx.collaboration.agentOrigin(execCtx.agentName ?? 'pptx-tool');
				await codec.hydrate(room.ydoc, outputBytes, origin);

				// Update viewer if available
				if (execCtx.viewer) {
					execCtx.viewer.replaceContent(filePath, outputBytes, {
						markDirty: false,
						mimeType: PPTX_MIME_TYPE,
					});
				}

				return { savedToDisk: false, routedThroughCollaboration: true };
			}
		}
	}

	// Non-collaboration: write to disk
	await execCtx.filesystem.writeFile(filePath, outputBytes);

	// Update viewer if available
	if (execCtx.viewer) {
		execCtx.viewer.replaceContent(filePath, outputBytes, {
			markDirty: false,
			mimeType: PPTX_MIME_TYPE,
		});
	}

	return { savedToDisk: true, routedThroughCollaboration: false };
}

/**
 * Execute a pure tool function with full collaboration-aware load/save pipeline.
 * Handles dehydrate-from-room → tool → hydrate-back-to-room routing automatically.
 */
export async function executeToolWithContext<T>(
	filePath: string,
	execCtx: ExecutionContext,
	toolFn: (ctx: ToolContext) => ToolResult<T> | Promise<ToolResult<T>>,
): Promise<T & { savedToDisk?: boolean; routedThroughCollaboration?: boolean }> {
	const { pptxData, rawBytes } = await loadPresentation(filePath, execCtx);

	const toolResult = await toolFn({ pptxData });

	if (toolResult.dirty) {
		const saveResult = await savePresentation(
			filePath,
			toolResult.pptxData,
			rawBytes,
			execCtx,
			toolResult.saveOptions,
		);
		return { ...toolResult.result, ...saveResult };
	}

	return toolResult.result as T & {
		savedToDisk?: boolean;
		routedThroughCollaboration?: boolean;
	};
}
