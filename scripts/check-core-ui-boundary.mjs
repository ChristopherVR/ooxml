/** Core must never depend on the UI. Usage: node scripts/check-core-ui-boundary.mjs [root] */
import { readdirSync, readFileSync } from 'node:fs';
import { dirname, join, relative, resolve, sep } from 'node:path';
import { fileURLToPath } from 'node:url';

const ROOT = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const DEPENDENCY_FIELDS = [
	'dependencies',
	'devDependencies',
	'peerDependencies',
	'optionalDependencies',
];
const isUi = (specifier) => specifier === 'ooxml-ui' || specifier.startsWith('ooxml-ui/');

function* sourceFiles(directory) {
	for (const entry of readdirSync(directory, { withFileTypes: true })) {
		const path = join(directory, entry.name);
		if (entry.isDirectory()) yield* sourceFiles(path);
		else if (/\.(?:[cm]?[jt]s|[jt]sx)$/.test(entry.name) && !entry.name.includes('.test.')) {
			yield path;
		}
	}
}

/**
 * Module specifiers of a source file, found with a small tokenizer: strings, templates, regular
 * expressions and comments are skipped, so only real `import`, `export ... from`, `import(...)`,
 * `require(...)` and `require.resolve(...)` positions count. (The TypeScript compiler API is not
 * available from TypeScript 7, and a boundary check should not need a compiler anyway.)
 */
function tokens(text) {
	const out = [];
	let i = 0;
	let line = 1;
	let lineStart = 0;
	const newline = (at) => {
		line += 1;
		lineStart = at + 1;
	};
	const previous = () => out.at(-1);
	const regexAllowed = () => {
		const last = previous();
		if (!last) return true;
		if (last.type === 'word')
			return /^(?:return|typeof|case|in|of|delete|void|throw|new|else|do)$/.test(last.text);
		return last.type === 'punct' && !/^[)\]}]$/.test(last.text);
	};
	while (i < text.length) {
		const ch = text[i];
		if (ch === '\n') {
			newline(i);
			i += 1;
		} else if (/\s/.test(ch)) i += 1;
		else if (ch === '/' && text[i + 1] === '/') {
			while (i < text.length && text[i] !== '\n') i += 1;
		} else if (ch === '/' && text[i + 1] === '*') {
			i += 2;
			while (i < text.length && !(text[i] === '*' && text[i + 1] === '/')) {
				if (text[i] === '\n') newline(i);
				i += 1;
			}
			i += 2;
		} else if (ch === '"' || ch === "'" || ch === '`') {
			const start = i;
			const startLine = line;
			const startColumn = i - lineStart + 1;
			let value = '';
			let plain = true;
			i += 1;
			while (i < text.length && text[i] !== ch) {
				if (text[i] === '\\') {
					value += text[i + 1] ?? '';
					i += 2;
					continue;
				}
				if (ch === '`' && text[i] === '$' && text[i + 1] === '{') plain = false;
				if (text[i] === '\n') newline(i);
				value += text[i];
				i += 1;
			}
			i += 1;
			out.push({ type: 'string', text: value, plain, line: startLine, column: startColumn, start });
		} else if (ch === '/' && regexAllowed()) {
			i += 1;
			let inClass = false;
			while (i < text.length && (text[i] !== '/' || inClass) && text[i] !== '\n') {
				if (text[i] === '\\') i += 1;
				else if (text[i] === '[') inClass = true;
				else if (text[i] === ']') inClass = false;
				i += 1;
			}
			i += 1;
			out.push({ type: 'regex', text: '' });
		} else if (/[\w$]/.test(ch)) {
			const start = i;
			while (i < text.length && /[\w$]/.test(text[i])) i += 1;
			out.push({ type: 'word', text: text.slice(start, i) });
		} else {
			out.push({ type: 'punct', text: ch });
			i += 1;
		}
	}
	return out;
}

function importSpecifiers(text) {
	const found = [];
	const list = tokens(text);
	const word = (token, value) => token?.type === 'word' && token.text === value;
	const punct = (token, value) => token?.type === 'punct' && token.text === value;
	list.forEach((token, index) => {
		if (token.type !== 'string' || !token.plain || !isUi(token.text)) return;
		const [one, two, three, four] = [1, 2, 3, 4].map((back) => list[index - back]);
		const isSpecifier =
			word(one, 'from') ||
			word(one, 'import') ||
			(punct(one, '(') && (word(two, 'import') || word(two, 'require'))) ||
			(punct(one, '(') && word(two, 'resolve') && punct(three, '.') && word(four, 'require'));
		if (isSpecifier) found.push(token);
	});
	return found;
}

/** `/// <reference types="..." />` directives, which live in comments. */
function referenceTypes(text) {
	return [...text.matchAll(/^\s*\/\/\/\s*<reference\s+types\s*=\s*["']([^"']+)["']/gm)]
		.map((match) => match[1])
		.filter(isUi);
}

export function checkCoreUiBoundary(root = ROOT) {
	const problems = [];
	const manifest = JSON.parse(readFileSync(join(root, 'package.json'), 'utf8'));
	for (const field of DEPENDENCY_FIELDS) {
		for (const [name, version] of Object.entries(manifest[field] ?? {})) {
			if (isUi(name) || /^npm:ooxml-ui(?:@|\/|$)/.test(String(version))) {
				problems.push(`package.json: ${field}.${name} depends on ooxml-ui`);
			}
		}
	}
	for (const path of sourceFiles(join(root, 'src'))) {
		// src/ui is ooxml-ui itself, a separate package that sits in this tree.
		if (relative(root, path).split(sep).slice(0, 2).join('/') === 'src/ui') continue;
		const text = readFileSync(path, 'utf8');
		for (const token of importSpecifiers(text))
			problems.push(`${relative(root, path)}:${token.line}:${token.column}: imports ${token.text}`);
		for (const name of referenceTypes(text))
			problems.push(`${relative(root, path)}: references types from ${name}`);
	}
	return problems;
}

if (process.argv[1] && resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
	const problems = checkCoreUiBoundary(process.argv[2] ? resolve(process.argv[2]) : ROOT);
	if (problems.length) {
		console.error(
			'src/ (the core) must not depend on ooxml-ui; the UI depends on the core, never the reverse.',
		);
		for (const problem of problems) console.error(problem);
		process.exitCode = 1;
	} else console.log('Core/UI dependency boundary passed.');
}
