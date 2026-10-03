/** Core must never depend on the UI. Usage: node scripts/check-core-ui-boundary.mjs [root] */
import { readdirSync, readFileSync } from 'node:fs';
import { dirname, join, relative, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import ts from 'typescript';

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

function importSpecifiers(source) {
	const specifiers = [];
	const add = (node) => {
		if (node && ts.isStringLiteralLike(node) && isUi(node.text)) specifiers.push(node);
	};
	const visit = (node) => {
		if (ts.isImportDeclaration(node) || ts.isExportDeclaration(node)) add(node.moduleSpecifier);
		else if (
			ts.isImportEqualsDeclaration(node) &&
			ts.isExternalModuleReference(node.moduleReference)
		) {
			add(node.moduleReference.expression);
		} else if (ts.isImportTypeNode(node) && ts.isLiteralTypeNode(node.argument)) {
			add(node.argument.literal);
		} else if (ts.isCallExpression(node)) {
			const callee = node.expression;
			if (
				callee.kind === ts.SyntaxKind.ImportKeyword ||
				(ts.isIdentifier(callee) && callee.text === 'require') ||
				(ts.isPropertyAccessExpression(callee) &&
					ts.isIdentifier(callee.expression) &&
					callee.expression.text === 'require' &&
					callee.name.text === 'resolve')
			)
				add(node.arguments[0]);
		}
		ts.forEachChild(node, visit);
	};
	visit(source);
	return specifiers;
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
		const source = ts.createSourceFile(
			path,
			readFileSync(path, 'utf8'),
			ts.ScriptTarget.Latest,
			true,
		);
		for (const node of importSpecifiers(source)) {
			const position = source.getLineAndCharacterOfPosition(node.getStart(source));
			problems.push(
				`${relative(root, path)}:${position.line + 1}:${position.character + 1}: imports ${node.text}`,
			);
		}
		for (const reference of source.typeReferenceDirectives) {
			if (isUi(reference.fileName))
				problems.push(`${relative(root, path)}: references types from ${reference.fileName}`);
		}
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
