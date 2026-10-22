import { globSync, readdirSync, readFileSync } from 'node:fs';
import path from 'node:path';

import { describe, expect, it } from 'vitest';

const SRC = path.resolve(import.meta.dirname, '../..');

/** The only source modules allowed to import react-i18next. */
const HOST_SOURCE_MODULES = ['viewer/i18n/react-i18next-source.ts'];

/** `... from 'x'`, `import('x')` and a bare `import 'x'`. */
const IMPORT_FROM =
	/\bfrom\s+['"]([^'"]+)['"]|\bimport\s*\(\s*['"]([^'"]+)['"]\s*\)|\bimport\s+['"]([^'"]+)['"]/g;

const NAMED_USE_TRANSLATION =
	/import\s*\{[^}]*\buseTranslation\b[^}]*\}\s*from\s*['"]([^'"]+)['"]/g;

const isSource = (name: string) => /\.tsx?$/.test(name) && !/\.test\.tsx?$/.test(name);

function sourceFiles(dir: string): string[] {
	return readdirSync(dir, { withFileTypes: true }).flatMap((entry) => {
		const full = path.join(dir, entry.name);
		if (entry.isDirectory()) {
			return sourceFiles(full);
		}
		return isSource(entry.name) ? [full] : [];
	});
}

function parse(file: string): { name: string; source: string; imports: string[] } {
	// latin1 keeps every byte: one module carries a literal NUL that utf8 tools treat as binary.
	const source = readFileSync(file, 'latin1');
	return {
		name: path.relative(SRC, file).split(path.sep).join('/'),
		source,
		imports: [...source.matchAll(IMPORT_FROM)].map(
			(match) => match[1] ?? match[2] ?? match[3] ?? '',
		),
	};
}

describe('react-i18next boundary', () => {
	const files = sourceFiles(SRC).map(parse);

	it('scans the whole package source', () => {
		const expected = globSync('**/*.{ts,tsx}', { cwd: SRC }).filter(isSource);
		expect(expected.length).toBeGreaterThan(0);
		expect(files).toHaveLength(expected.length);
	});

	it('matches every import form', () => {
		const imports = parse(path.join(SRC, 'viewer/i18n/react-i18next-source.ts')).imports;
		expect(imports).toContain('react-i18next');
		const forms = "import 'a';\nimport b from 'b';\nexport { c } from 'c';\nawait import('d');";
		expect([...forms.matchAll(IMPORT_FROM)].map((m) => m[1] ?? m[2] ?? m[3])).toStrictEqual([
			'a',
			'b',
			'c',
			'd',
		]);
	});

	it('imports react-i18next only from the host source module', () => {
		const importers = files
			.filter((file) => file.imports.some((spec) => /^(?:react-)?i18next(?:\/|$)/.test(spec)))
			.map((file) => file.name);
		expect(importers.toSorted()).toStrictEqual(HOST_SOURCE_MODULES);
	});

	it('routes every useTranslation import through the internal leaf hook', () => {
		const offenders = files
			.filter((file) =>
				[...file.source.matchAll(NAMED_USE_TRANSLATION)].some(
					(match) => !/(?:^|\/)use-translation$/.test(match[1] ?? ''),
				),
			)
			.map((file) => file.name)
			.filter((name) => !HOST_SOURCE_MODULES.includes(name));
		expect(offenders).toStrictEqual([]);
	});
});
