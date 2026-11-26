// Test-only: every element this package defines, found from its public register functions and
// from every `definer` in the source, so the conformance suite covers a new element without a
// list to keep. Modules are imported by computed path so the strict project does not type-check
// the PowerPoint sources (see tsconfig.json).
import { readdirSync, readFileSync } from 'node:fs';
import { join, relative, resolve, sep } from 'node:path';

export type ElementClass = CustomElementConstructor & {
	elementProperties?: Map<PropertyKey, { type?: unknown; state?: boolean; attribute?: unknown }>;
	elementStyles?: { cssText: string }[];
	observedAttributes?: string[];
	watched?: readonly string[];
	syncUpdates?: boolean;
};

export interface DefinedElement {
	tag: string;
	ctor: ElementClass;
	/** False when only a `definer` the public entries never call defines it. */
	public: boolean;
}

/** The public register functions, by the module that exports them (paths under `src/`). */
export const ENTRIES: readonly { module: string; register: readonly string[] }[] = [
	{
		module: 'index.ts',
		register: ['registerOfficeUi', 'defineOfficeAddInsDialog', 'registerInkElements'],
	},
	{
		module: 'teams/index.ts',
		register: [
			'registerTeams',
			'defineTeamsApp',
			'defineTeamsSettings',
			'defineTeamsChannelTab',
			'defineTeamsFilesPanel',
			'defineTeamsPresentationPreview',
			'defineTeamsContentPreview',
		],
	},
	{ module: 'pptx/index.ts', register: ['registerPptxWebControls', 'defineThreeViewElement'] },
	{ module: 'docx/index.ts', register: ['registerDocxEditor'] },
	{ module: 'xlsx/index.ts', register: ['defineXlsxEditor'] },
	{ module: 'visio/index.ts', register: ['registerVisioViewer'] },
];

const SRC = resolve(import.meta.dirname, '..');
const SKIP = /\.test\.ts$|\.d\.ts$|[\\/](__fixtures__|conformance)[\\/]/;

/** Non-test source files (relative to `src/`, with `/`) whose text matches `pattern`. */
export function sourcesMatching(pattern: RegExp, dir = SRC): string[] {
	return readdirSync(dir, { withFileTypes: true }).flatMap((entry) => {
		const path = join(dir, entry.name);
		if (entry.isDirectory()) return sourcesMatching(pattern, path);
		const name = relative(SRC, path).split(sep).join('/');
		if (!name.endsWith('.ts') || SKIP.test(`/${name}`)) return [];
		return pattern.test(readFileSync(path, 'utf8')) ? [name] : [];
	});
}

const load = (module: string): Promise<Record<string, unknown>> =>
	import(/* @vite-ignore */ `../${module}`);

type Define = CustomElementRegistry['define'];

/**
 * Calls every public register function twice (a second call must be a no-op), then every
 * `definer` export in the source, and reports each element with the class it registered.
 */
export async function defineEverything(registry: CustomElementRegistry): Promise<{
	elements: DefinedElement[];
	errors: string[];
}> {
	const defined = new Map<string, DefinedElement>();
	let reached = true;
	const define: Define = registry.define;
	registry.define = function (this: CustomElementRegistry, tag, ctor, options) {
		defined.set(tag, { tag, ctor: ctor as ElementClass, public: reached });
		return define.call(this, tag, ctor, options);
	};
	const errors: string[] = [];
	const call = async (name: string, run: () => unknown) => {
		try {
			await run();
		} catch (error) {
			errors.push(`${name}: ${String(error)}`);
		}
	};
	try {
		for (const { module, register } of ENTRIES) {
			const exports = await load(module);
			for (const name of register)
				for (const round of [1, 2])
					await call(`${module} ${name} (call ${round})`, () => (exports[name] as () => void)());
		}
		reached = false;
		for (const module of sourcesMatching(/\bdefiner\(/)) {
			const exports = await load(module);
			for (const [name, value] of Object.entries(exports))
				if (typeof value === 'function' && typeof (value as { tag?: unknown }).tag === 'string')
					await call(`${module} ${name}`, () => (value as () => void)());
		}
	} finally {
		registry.define = define;
	}
	return { elements: [...defined.values()].sort((a, b) => a.tag.localeCompare(b.tag)), errors };
}
