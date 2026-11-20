/**
 * The seam that lets the in-viewer AI assistant run the SAME tool logic the
 * `pptx-viewer-mcp` server exposes, against the LIVE open deck instead of a file
 * on disk.
 *
 * Each MCP tool is a pure(ish) `PptxData` transform: `fn(ctx, params) ->
 * { pptxData, result, dirty }`. Here we assemble a working {@link PptxData} from
 * the bridge's live slides + theme (overlaid on the binding's full parsed data
 * when it exposes {@link PptxAiBridge.getDeckData}), run the transform, and send
 * every write through {@link routeDeckWrite}, so the active write policy decides
 * whether it is applied or held for review. The commit strategy only names the
 * bridge write that lands it as one Ctrl+Z:
 *
 * - `slides` -> {@link PptxAiBridge.applySlidesUpdate}.
 * - `theme`  -> {@link PptxAiBridge.applyTheme}.
 * - `deck`   -> {@link PptxAiBridge.applyDeckData} for presentation-level state
 *   (metadata, sections, canvas size, presentation properties); reported as
 *   unavailable when the binding does not implement that optional seam.
 * - `read`   -> return the tool's result value, no commit.
 */

import type { ToolContext, ToolResult } from 'ooxml-core/pptx/automation';

import { readLiveDeck } from '../proposals';
import type { AiToolContext } from './executor-base';
import { routeDeckWrite } from './executor-base';

/** How a shared tool's output is committed to the live deck. */
export type SharedToolCommit = 'read' | 'slides' | 'theme' | 'deck';

/** An MCP tool function: a `PptxData` transform. Params exclude `filePath`. */
export type SharedToolFn = (
	ctx: ToolContext,
	params: never,
) => ToolResult<unknown> | Promise<ToolResult<unknown>>;

/** Registry metadata needed to run and commit one shared tool. */
export interface SharedToolSpec {
	/** The `pptx-viewer-mcp` tool function. */
	fn: SharedToolFn;
	/** Commit strategy for the tool's `PptxData` output. */
	commit: SharedToolCommit;
	/** History-entry / proposal label used when the tool writes. */
	label: string;
	/** Force explicit approval before applying (e.g. slide deletion). */
	forceApproval?: boolean;
}

/** Spread a tool result value into an object so we can annotate it with write status. */
function asObject(value: unknown): Record<string, unknown> {
	return value && typeof value === 'object'
		? (value as Record<string, unknown>)
		: { result: value };
}

/**
 * Run one `pptx-viewer-mcp` tool on a private copy of the live deck and route
 * any change through the write policy. Returns a JSON-serialisable value for
 * the model: the tool's own result for reads, plus write-routing status
 * (`staged` / `applied` / `proposalId` / `summary`) for writes.
 */
export async function runSharedTool(
	ctx: AiToolContext,
	spec: SharedToolSpec,
	input: unknown,
): Promise<unknown> {
	const before = structuredClone(readLiveDeck(ctx.bridge));
	const toolCtx: ToolContext = { pptxData: structuredClone(before) };
	const res = await spec.fn(toolCtx, input as never);

	if (spec.commit === 'read' || !res.dirty) {
		return res.result;
	}
	if (spec.commit === 'deck' && !ctx.bridge.applyDeckData) {
		throw new Error(
			'This change affects presentation-level data (metadata, sections, layout, or canvas size), which this viewer cannot apply yet.',
		);
	}
	const write = routeDeckWrite(ctx, {
		label: spec.label,
		commit: spec.commit,
		before,
		after: res.pptxData,
		forceApproval: spec.forceApproval,
	});
	return { ...asObject(res.result), ...write };
}
