import type { z } from 'zod';

import { editTool, readTool } from '../../../automation/tools/define';
import type { AutomationTool } from '../../../automation/tools/types';
import { PptxHandler } from '../../index';
import { arrayBufferOf, saveDeckBytes } from '../execution';
import type { ToolContext, ToolResult } from '../types';

/** Decks the tools read: PowerPoint files and legacy `.ppt`, which the core converts on load. */
export const DECK_READ_EXTENSIONS: readonly string[] = ['.pptx', '.ppt'];
/** Edits save PresentationML, so they only write `.pptx`. */
export const DECK_EDIT_EXTENSIONS: readonly string[] = ['.pptx'];

type FileShape = z.ZodRawShape & { filePath: z.ZodTypeAny };
type DeckParams<S extends FileShape> = z.infer<z.ZodObject<Omit<S, 'filePath'>>>;

/**
 * A PowerPoint tool over bytes: loads the deck, runs a core tool function on it and, for an
 * edit that changed the deck, returns the saved bytes. The schema's `filePath` (kept for the
 * assistant schemas that still name the file) is dropped; the transport supplies the file.
 */
export function deckTool<S extends FileShape>(
	kind: 'read' | 'edit',
	name: string,
	description: string,
	schema: z.ZodObject<S>,
	fn: (ctx: ToolContext, params: DeckParams<S>) => ToolResult | Promise<ToolResult>,
	options: { returnsText?: boolean } = {},
): AutomationTool {
	const { filePath: _filePath, ...input } = schema.shape;
	const definition = { name, description, format: 'pptx' as const, input };
	const run = async (bytes: Uint8Array, params: DeckParams<S>) => {
		const pptxData = await new PptxHandler().load(arrayBufferOf(bytes));
		return fn({ pptxData }, params);
	};
	if (kind === 'read')
		return readTool({
			...definition,
			extensions: DECK_READ_EXTENSIONS,
			run: async (bytes, params) => {
				const result = await run(bytes, params as DeckParams<S>);
				if (result.dirty) throw new Error(`${name} changed the deck but is a read tool`);
				return result.result;
			},
		});
	return editTool({
		...definition,
		extensions: DECK_EDIT_EXTENSIONS,
		...(options.returnsText && { returnsText: true }),
		run: async (bytes, params) => {
			const result = await run(bytes, params as DeckParams<S>);
			const report = reportOf(result.result);
			return result.dirty ? { ...report, bytes: await saveDeckBytes(result, bytes) } : report;
		},
	});
}

/** A tool's result as the report object: objects as they are, anything else under `result`. */
function reportOf(value: unknown): Record<string, unknown> {
	return value && typeof value === 'object' && !Array.isArray(value)
		? (value as Record<string, unknown>)
		: { result: value };
}
