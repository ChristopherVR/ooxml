import { z } from 'zod';
import {
	UNTRUSTED_NOTE,
	type AutomationCreateTool,
	type AutomationEditResult,
	type AutomationEditTool,
	type AutomationFormat,
	type AutomationReadTool,
} from './types';

type Params<S extends z.ZodRawShape> = z.infer<z.ZodObject<S>>;

interface Definition<S extends z.ZodRawShape> {
	name: string;
	description: string;
	/** Schema version; omit for 1. */
	version?: number;
	format: AutomationFormat;
	extensions: readonly string[];
	input: S;
}

const path = z.string().min(1);

function base<S extends z.ZodRawShape>(
	definition: Definition<S>,
	kind: 'read' | 'create' | 'edit',
) {
	const schema = z.strictObject(definition.input);
	let json: Record<string, unknown> | undefined;
	return {
		parse: (params: unknown): Params<S> => {
			const parsed = schema.safeParse(params ?? {});
			if (!parsed.success)
				throw new Error(`Invalid ${definition.name} parameters: ${z.prettifyError(parsed.error)}`);
			return parsed.data;
		},
		tool: {
			name: definition.name,
			description: definition.description,
			version: definition.version ?? 1,
			format: definition.format,
			extensions: definition.extensions,
			input: definition.input,
			fileInput: {
				filePath: path,
				...(kind === 'edit' && { outputPath: path.optional() }),
				...definition.input,
			},
			jsonSchema: () => (json ??= z.toJSONSchema(schema) as Record<string, unknown>),
		},
	};
}

export function readTool<S extends z.ZodRawShape>(
	definition: Definition<S> & { run(bytes: Uint8Array, params: Params<S>): Promise<unknown> },
): AutomationReadTool {
	const { parse, tool } = base(definition, 'read');
	return {
		...tool,
		description: definition.description + UNTRUSTED_NOTE,
		kind: 'read',
		run: async (bytes, params) => definition.run(bytes, parse(params)),
	};
}

export function createTool<S extends z.ZodRawShape>(
	definition: Definition<S> & { run(params: Params<S>): Promise<Uint8Array> },
): AutomationCreateTool {
	const { parse, tool } = base(definition, 'create');
	return { ...tool, kind: 'create', run: async (params) => definition.run(parse(params)) };
}

export function editTool<S extends z.ZodRawShape>(
	definition: Definition<S> & {
		run(bytes: Uint8Array, params: Params<S>): Promise<AutomationEditResult>;
		/** The report can carry document text (a list action), so it gets the untrusted note. */
		returnsText?: boolean;
	},
): AutomationEditTool {
	const { parse, tool } = base(definition, 'edit');
	return {
		...tool,
		description: definition.description + (definition.returnsText ? UNTRUSTED_NOTE : ''),
		kind: 'edit',
		run: async (bytes, params) => definition.run(bytes, parse(params)),
	};
}

/** Shared parameter pieces. */
export const id = path;
export const index = z.number().int().min(0);
export const longText = z.string().max(1_000_000);
