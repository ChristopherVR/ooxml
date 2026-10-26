export interface BareImport {
	readonly specifier: string;
	/** esbuild's import kind: `import-statement`, `dynamic-import`, `require-call`, ... */
	readonly kind: string;
	readonly importer: string;
}

export interface ModuleGraph {
	/** Every module reached through relative imports, the entry first. */
	readonly files: string[];
	/** Relative `import()` targets not walked (only with `followDynamic: false`). */
	readonly lazy: string[];
	/** Every bare specifier the reached modules import at runtime. */
	readonly bare: BareImport[];
}

export function walkModuleGraph(
	entry: string,
	options?: { followDynamic?: boolean },
): Promise<ModuleGraph>;
