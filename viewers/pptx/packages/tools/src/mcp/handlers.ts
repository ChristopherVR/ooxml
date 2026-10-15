import { isAbsolute, resolve, sep } from 'node:path';

import { createFileOperations } from 'ooxml-core/automation/node';

import { executeToolWithContext } from '../execution';
import type { ExecutionContext, ToolContext, ToolResult } from '../types';

const ALLOWED_EXTENSIONS = new Set(['.pptx', '.ppt']);

/**
 * Resolve `dir` against `rootDir` and ensure the result stays within it.
 *
 * `rootDir` defaults to `process.env.PPTX_TOOLS_ROOT` if set, else
 * `process.cwd()`. Throws on traversal escape or non-string input. Node-only:
 * this lives here (rather than beside the tool it guards) so the tool modules
 * stay free of `node:path` and remain bundleable for the browser.
 */
export function resolveScopedDir(dir: string, rootDir?: string): string {
	if (typeof dir !== 'string' || dir.length === 0) {
		throw new Error('MCP tool: path must be a non-empty string');
	}

	const root = resolve(rootDir ?? process.env['PPTX_TOOLS_ROOT'] ?? process.cwd());
	const resolved = isAbsolute(dir) ? resolve(dir) : resolve(root, dir);

	const rootWithSep = root.endsWith(sep) ? root : root + sep;
	if (resolved !== root && !resolved.startsWith(rootWithSep)) {
		throw new Error(`MCP tool: path "${dir}" resolves outside the allowed root`);
	}

	return resolved;
}

/**
 * Resolve `filePath` against `rootDir` and ensure the result stays within it.
 *
 * - `rootDir` defaults to `process.env.PPTX_TOOLS_ROOT` if set, else `process.cwd()`.
 * - The path must end in `.pptx` or `.ppt` (case-insensitive).
 * - Throws on traversal escape, non-string inputs, or disallowed extensions.
 */
export function resolveScopedFilePath(filePath: string, rootDir?: string): string {
	if (typeof filePath !== 'string' || filePath.length === 0) {
		throw new Error('MCP tool: filePath must be a non-empty string');
	}

	const resolved = resolveScopedDir(filePath, rootDir);

	const lower = resolved.toLowerCase();
	let isAllowed = false;
	for (const ext of ALLOWED_EXTENSIONS) {
		if (lower.endsWith(ext)) {
			isAllowed = true;
			break;
		}
	}
	if (!isAllowed) {
		throw new Error(
			`MCP tool: filePath "${filePath}" must end with one of: ${Array.from(ALLOWED_EXTENSIONS).join(', ')}`,
		);
	}

	return resolved;
}

/** Optional configuration for the MCP context. */
export interface McpContextOptions {
	/**
	 * Root directory under which all `filePath` arguments are scoped.
	 * Defaults to `process.env.PPTX_TOOLS_ROOT` if set, else `process.cwd()`.
	 */
	rootDir?: string;
}

let warnedCwdRoot = false;
const operations = new Map<string, ReturnType<typeof createFileOperations>>();

/**
 * File operations confined (by realpath, not just lexically) to the configured
 * root. Instances are shared per root so per-path serialisation holds across
 * tool calls.
 */
export function fileOperationsFor(options?: McpContextOptions) {
	let root = options?.rootDir ?? process.env['PPTX_TOOLS_ROOT'];
	if (!root) {
		if (!warnedCwdRoot) {
			warnedCwdRoot = true;
			process.stderr.write(
				'pptx-viewer-tools: no root configured (--root or PPTX_TOOLS_ROOT); using the current working directory\n',
			);
		}
		root = process.cwd();
	}
	const key = resolve(root);
	let files = operations.get(key);
	if (!files) {
		files = createFileOperations(key);
		operations.set(key, files);
	}
	return files;
}

const EXTENSIONS: readonly string[] = [...ALLOWED_EXTENSIONS];

/** Create an ExecutionContext for MCP (no collaboration, no viewer). */
export function createMcpContext(options?: McpContextOptions): ExecutionContext {
	const files = fileOperationsFor(options);
	return {
		filesystem: {
			readFile: (path: string) => files.readFile(path, EXTENSIONS),
			writeFile: (path: string, data: Uint8Array) => files.writeFile(path, EXTENSIONS, data),
		},
	};
}

/** Confine an output directory under the root; returns a root-relative path. */
export function confineOutputDir(dir: string, options?: McpContextOptions): Promise<string> {
	if (typeof dir !== 'string' || dir.length === 0) {
		return Promise.reject(new Error('MCP tool: path must be a non-empty string'));
	}
	return fileOperationsFor(options).directory(dir);
}

/** Read-modify-write a confined presentation file under its per-path lock. */
export function withMcpFile<T>(
	filePath: string,
	options: McpContextOptions | undefined,
	action: (bytes: Uint8Array) => Promise<{ value: T; bytes?: Uint8Array }>,
): Promise<T> {
	const files = fileOperationsFor(options);
	return files.exclusive(filePath, EXTENSIONS, async () => {
		const { value, bytes } = await action(await files.readFile(filePath, EXTENSIONS));
		if (bytes) await files.writeFile(filePath, EXTENSIONS, bytes);
		return value;
	});
}

/**
 * Run a tool via MCP, using the shared collaboration-aware execution pipeline.
 * Returns the inner result value (not the full ToolResult wrapper).
 *
 * `filePath` is confined by realpath to the configured root (option
 * `rootDir`, env `PPTX_TOOLS_ROOT`, or process.cwd() with a warning), size
 * capped, and written atomically under a per-path lock.
 */
export async function runMcpTool<T>(
	filePath: string,
	fn: (ctx: ToolContext) => ToolResult<T> | Promise<ToolResult<T>>,
	options?: McpContextOptions,
): Promise<T> {
	if (typeof filePath !== 'string' || filePath.length === 0) {
		throw new Error('MCP tool: filePath must be a non-empty string');
	}
	const ctx = createMcpContext(options);
	return fileOperationsFor(options).exclusive(filePath, EXTENSIONS, () =>
		executeToolWithContext(filePath, ctx, fn),
	);
}
