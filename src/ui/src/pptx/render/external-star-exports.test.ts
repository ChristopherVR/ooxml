import { readFileSync, readdirSync } from 'node:fs';
import { join, relative } from 'node:path';
import { fileURLToPath } from 'node:url';

import {
	isExportDeclaration,
	isNamespaceExport,
	isStringLiteral,
} from 'typescript/unstable/ast/is';
import { API } from 'typescript/unstable/sync';
import { expect, it } from 'vitest';

// The bundler turns `export * from '<package>'` inside ooxml-ui/pptx into a runtime namespace copy
// (`__reExport`) that reaches the package entry, and from it every export of ooxml-ui/pptx, so a
// host bundler can no longer drop any of them. Re-export external modules by name instead.
//
// The check parses every source with the TypeScript compiler the package builds with, so quotes,
// line breaks, indentation, a missing semicolon or a `.tsx` file cannot hide an offender, and
// comments or strings that merely contain the text are not reported. `export type *` is erased
// from the bundle and stays allowed.

const pptxRoot = fileURLToPath(new URL('..', import.meta.url));
const slash = (path: string): string => path.replaceAll('\\', '/');

function sources(dir: string): string[] {
	return readdirSync(dir, { withFileTypes: true }).flatMap((entry): string[] => {
		const path = join(dir, entry.name);
		if (entry.isDirectory()) return entry.name === 'node_modules' ? [] : sources(path);
		return /\.tsx?$/u.test(entry.name) && !/\.(test|d)\.tsx?$/u.test(entry.name) ? [path] : [];
	});
}

const isBareSpecifier = (specifier: string): boolean => !/^[./#]/u.test(specifier);

/**
 * Every value `export *` (including `export * as name`) of a bare package specifier in `files`,
 * as `<path relative to src/pptx>: <specifier>`. `virtual` supplies file contents by path, for
 * files that do not exist on disk.
 */
function externalStarExports(files: string[], virtual: Record<string, string> = {}): string[] {
	const config = join(pptxRoot, '__external-star-exports__', 'tsconfig.json');
	const contents = new Map(Object.entries(virtual).map(([path, text]) => [slash(path), text]));
	contents.set(
		slash(config),
		JSON.stringify({ compilerOptions: { noLib: true, noResolve: true, types: [] }, files }),
	);
	const api = new API({
		cwd: pptxRoot,
		fs: {
			readFile: (path) => contents.get(slash(path)),
			fileExists: (path) => (contents.has(slash(path)) ? true : undefined),
		},
	});
	try {
		const project = api.updateSnapshot({ openProjects: [config] }).getProject(config);
		if (!project) throw new Error(`no project for ${config}`);
		return files.flatMap((file) => {
			const source = project.program.getSourceFile(file);
			if (!source) throw new Error(`${file} was not parsed`);
			return source.statements.flatMap((statement) => {
				if (
					!isExportDeclaration(statement) ||
					statement.isTypeOnly ||
					!statement.moduleSpecifier ||
					!isStringLiteral(statement.moduleSpecifier) ||
					(statement.exportClause && !isNamespaceExport(statement.exportClause)) ||
					!isBareSpecifier(statement.moduleSpecifier.text)
				) {
					return [];
				}
				return [`${slash(relative(pptxRoot, file))}: ${statement.moduleSpecifier.text}`];
			});
		});
	} finally {
		api.close();
	}
}

/**
 * Whether `text` may hold an `export *`: the keyword followed by `*` with only whitespace or
 * comments between, as every such declaration has. Only these files go to the compiler: handing
 * it all of src/pptx took past two minutes on a loaded machine.
 */
const mayStarExport = (text: string): boolean =>
	/\bexport\b(?:\s|\/\*(?:[^*]|\*(?!\/))*\*\/|\/\/[^\r\n]*[\r\n])*\*/u.test(text);

// Compiles the candidates in a separate process, which a busy machine can slow past the default.
it('re-exports external modules by name, never with export *', { timeout: 120_000 }, () => {
	const files = sources(pptxRoot);
	expect(files.length).toBeGreaterThan(1000);
	const candidates = files.filter((file) => mayStarExport(readFileSync(file, 'utf8')));
	expect(candidates.length).toBeGreaterThan(0);
	expect(externalStarExports(candidates)).toEqual([]);
});

it('reports every spelling of an external value star re-export, and only those', () => {
	const fixture = (name: string): string => join(pptxRoot, '__external-star-exports__', name);
	const files: Record<string, string> = {
		// The two facades as they were before they re-exported by name.
		'text-style-scope.ts':
			"// Compatibility exports: paragraph and run editing belongs to PowerPoint core.\nexport * from 'ooxml-core/pptx/editor/render/text-style-scope';\n",
		'paragraph-geometry-overrides.ts':
			"export * from 'ooxml-core/pptx/editor/render/paragraph-geometry-overrides';\n",
		'double-quotes.ts': 'export * from "double-quoted";\n',
		'multi-line.ts': "export *\n\tfrom\n\t'multi-line';\n",
		'no-semicolon.ts': "export * from 'no-semicolon'\nexport const value = 1\n",
		'indented.ts': "\t  export * from 'indented';\n",
		'component.tsx': "export * from 'from-tsx';\nexport const View = () => <div />;\n",
		'namespace.ts': "export * as helpers from 'namespace-star';\n",
		'commented.ts': "export /* why */ * from 'commented';\n",
		'allowed.ts': [
			"export type * from 'type-only';",
			"export type * as Types from 'type-only-namespace';",
			"export * from './relative';",
			"export * from '#internal';",
			"export { named } from 'named-package';",
			"// export * from 'in-a-comment';",
			'export const text = "export * from \'in-a-string\'";',
		].join('\n'),
	};
	const virtual = Object.fromEntries(
		Object.entries(files).map(([name, text]) => [fixture(name), text]),
	);
	// The text filter keeps every file that holds one.
	expect(Object.values(files).every(mayStarExport)).toBe(true);
	expect(externalStarExports(Object.keys(virtual), virtual)).toEqual([
		'__external-star-exports__/text-style-scope.ts: ooxml-core/pptx/editor/render/text-style-scope',
		'__external-star-exports__/paragraph-geometry-overrides.ts: ooxml-core/pptx/editor/render/paragraph-geometry-overrides',
		'__external-star-exports__/double-quotes.ts: double-quoted',
		'__external-star-exports__/multi-line.ts: multi-line',
		'__external-star-exports__/no-semicolon.ts: no-semicolon',
		'__external-star-exports__/indented.ts: indented',
		'__external-star-exports__/component.tsx: from-tsx',
		'__external-star-exports__/namespace.ts: namespace-star',
		'__external-star-exports__/commented.ts: commented',
	]);
});
