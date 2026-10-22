import { sha256 } from '../digest/sha256';
import type { AutomationFormat, AutomationTool } from './tools/types';

/**
 * Runs the headless document tools against editors that are open on screen. A host's own
 * assistant (or the suite's) registers each mounted editor under an id it chooses, and every call
 * names the document it acts on: there is no "current document", so a chat that spans several
 * tabs can never edit the wrong one. The tools are the same definitions the MCP servers use;
 * this only adds the snapshot and load-back around them.
 */

/** What a mounted editor must offer. The editor elements fit it through `liveEditorFromElement`. */
export interface LiveEditor {
	readonly format: AutomationFormat;
	/** The open document as a file, including unsaved edits. */
	save(): Promise<Uint8Array> | Uint8Array;
	/** Replaces the open document with `bytes`. */
	load(bytes: Uint8Array): Promise<void> | void;
	/** True while the editor refuses edits (a read-only role, a locked document). */
	readonly readOnly?: boolean;
}

/** A tool call: the target document, then the tool's own parameters. */
export interface LiveToolCall {
	/** The id the editor was registered under. Required: there is no default document. */
	target: string;
	params?: unknown;
}

export interface LiveToolResult {
	target: string;
	tool: string;
	/** What the tool returned: a read's view, an edit's report. */
	result: unknown;
	/** True when an edit changed the open document. */
	changed: boolean;
}

/** What an edit is about to do, for a host that wants to confirm it first. */
export interface LiveEditProposal {
	target: string;
	tool: string;
	params: unknown;
	/** The tool's report for the edit (warnings, results). */
	result: unknown;
	/** The file before and after, for a diff. */
	before: Uint8Array;
	after: Uint8Array;
}

export interface LiveToolHostOptions {
	/**
	 * Called with each edit before it reaches the editor. Return false to discard it. Leave it out
	 * to apply edits straight away.
	 */
	review?(proposal: LiveEditProposal): boolean | Promise<boolean>;
}

export interface LiveToolDescriptor {
	name: string;
	description: string;
	kind: 'read' | 'edit';
	format: AutomationFormat;
	/** JSON Schema of the parameters, with the required `target` added. */
	inputSchema: Record<string, unknown>;
}

export interface LiveToolHost {
	/** Registers an editor under `id`. Returns a function that removes it. */
	register(id: string, editor: LiveEditor): () => void;
	/** The registered documents, so an assistant can offer them to the model. */
	targets(): { id: string; format: AutomationFormat; readOnly: boolean }[];
	/** The read and edit tools that apply to the registered documents, as function-calling specs. */
	describe(): LiveToolDescriptor[];
	call(name: string, call: LiveToolCall): Promise<LiveToolResult>;
}

const digest = (bytes: Uint8Array) =>
	Array.from(sha256(bytes), (byte) => byte.toString(16).padStart(2, '0')).join('');

/** `tools` is the catalogue to serve: `automationTools.docx`, `pptxTools` and so on, in any mix. */
export function createLiveToolHost(
	tools: readonly AutomationTool[],
	options: LiveToolHostOptions = {},
): LiveToolHost {
	const editors = new Map<string, LiveEditor>();
	const serve = tools.filter((tool) => tool.kind !== 'create');
	// One edit at a time per document: two overlapping calls would both start from the same bytes.
	const queues = new Map<string, Promise<unknown>>();

	function editorFor(target: unknown): LiveEditor {
		if (typeof target !== 'string' || target === '')
			throw new Error('Every call needs a target: the id of the document to act on.');
		const editor = editors.get(target);
		if (!editor) throw new Error(`No open document is registered as "${target}".`);
		return editor;
	}

	async function run(name: string, call: LiveToolCall): Promise<LiveToolResult> {
		const tool = serve.find((candidate) => candidate.name === name);
		if (!tool) throw new Error(`Unknown tool "${name}".`);
		const editor = editorFor(call?.target);
		const target = call.target;
		if (tool.format !== editor.format)
			throw new Error(
				`${name} works on ${tool.format} documents, but "${target}" is a ${editor.format} document.`,
			);
		const before = new Uint8Array(await editor.save());
		if (tool.kind === 'read')
			return { target, tool: name, result: await tool.run(before, call.params), changed: false };
		if (editor.readOnly) throw new Error(`"${target}" is read-only, so it cannot be edited.`);
		const { bytes: after, ...result } = await (
			tool as Extract<AutomationTool, { kind: 'edit' }>
		).run(before, call.params);
		if (!after) return { target, tool: name, result, changed: false };
		if (options.review) {
			const accepted = await options.review({
				target,
				tool: name,
				params: call.params,
				result,
				before,
				after,
			});
			if (!accepted)
				return { target, tool: name, result: { ...result, discarded: true }, changed: false };
		}
		// The user may have typed while the tool ran or the host reviewed; loading `after` would
		// silently drop that, so refuse and let the caller read the document again.
		if (editor.readOnly)
			throw new Error(`"${target}" became read-only, so the edit was not applied.`);
		if (digest(new Uint8Array(await editor.save())) !== digest(before))
			throw new Error(`"${target}" changed while the edit was prepared; read it again and retry.`);
		await editor.load(after);
		return { target, tool: name, result, changed: true };
	}

	return {
		register(id, editor) {
			if (editors.has(id)) throw new Error(`A document is already registered as "${id}".`);
			editors.set(id, editor);
			return () => {
				if (editors.get(id) === editor) editors.delete(id);
			};
		},
		targets: () =>
			[...editors].map(([id, editor]) => ({
				id,
				format: editor.format,
				readOnly: Boolean(editor.readOnly),
			})),
		describe() {
			const open = new Set([...editors.values()].map((editor) => editor.format));
			return serve
				.filter((tool) => open.has(tool.format))
				.map((tool) => {
					const schema = tool.jsonSchema() as {
						properties?: Record<string, unknown>;
						required?: string[];
					};
					return {
						name: tool.name,
						description: tool.description,
						kind: tool.kind as 'read' | 'edit',
						format: tool.format,
						inputSchema: {
							type: 'object',
							additionalProperties: false,
							properties: {
								target: {
									type: 'string',
									description: 'The id of the open document to act on.',
								},
								params: schema,
							},
							required: ['target'],
						},
					};
				});
		},
		call(name, call) {
			const key = typeof call?.target === 'string' ? call.target : '';
			const next = (queues.get(key) ?? Promise.resolve()).then(
				() => run(name, call),
				() => run(name, call),
			);
			queues.set(
				key,
				next.catch(() => undefined),
			);
			return next;
		},
	};
}

/** The members of an editor element that `liveEditorFromElement` uses. */
export interface LiveEditorElement {
	saveBytes?(): Promise<Uint8Array> | Uint8Array;
	exportVsdx?(): Promise<{ bytes: Uint8Array }>;
	load(bytes: Uint8Array, name?: string): Promise<unknown> | unknown;
	commitEdit?(): boolean;
	readOnly?: boolean;
}

/**
 * Adapts a mounted editor element (`<docx-editor>`, `<xlsx-editor>`, `<pptx-editor>`,
 * `<visio-editor>`) so it can be registered with a host. `name` is the file name the editor shows.
 */
export function liveEditorFromElement(
	format: AutomationFormat,
	element: LiveEditorElement,
	name?: string,
): LiveEditor {
	return {
		format,
		save: async () => {
			if (element.commitEdit && !element.commitEdit())
				throw new Error('Finish the current cell edit before using the document.');
			if (format === 'visio' && element.exportVsdx) return (await element.exportVsdx()).bytes;
			if (!element.saveBytes) throw new Error('The editor cannot save its document.');
			return element.saveBytes();
		},
		load: async (bytes) => {
			await element.load(bytes, name);
		},
		get readOnly() {
			return element.readOnly === true;
		},
	};
}
