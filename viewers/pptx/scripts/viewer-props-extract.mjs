/**
 * Reads the public props/inputs/options and events of each PowerPoint binding from source with
 * the TypeScript compiler API. Used by check-viewer-props.mjs; it knows nothing of the manifest.
 *
 * Where each binding declares its surface:
 *  - react:   `PowerPointViewerProps` (callbacks are `onX` members)
 *  - vue:     `PowerPointViewerProps` plus the call signatures of `PowerPointViewerEmits`
 *  - angular: `input()` and `output()` members of `PowerPointViewerComponent` (an `alias`
 *             option names the input; callback inputs are `onX`)
 *  - svelte:  `PowerPointViewerProps` (callbacks are lowercase `onx` members)
 *  - vanilla: `PptxViewerOptions`, which extends `PptxViewerCallbacks`
 * Interfaces that extend another one (`ViewportFitOptions`, `PptxViewerCallbacks`) contribute the
 * parent's members.
 */
import { readFileSync } from 'node:fs';
import { createRequire } from 'node:module';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

const here = dirname(fileURLToPath(import.meta.url));
const root = join(here, '..');
// `typescript` is a dependency of this workspace root, so resolve it from here.
const ts = createRequire(join(root, 'package.json'))('typescript');

const FIT_OPTIONS = join(root, '..', '..', 'src', 'ui', 'src', 'pptx', 'render', 'viewport-fit.ts');

const sources = new Map();
function parse(file) {
	let sf = sources.get(file);
	if (!sf) {
		sf = ts.createSourceFile(file, readFileSync(file, 'utf8'), ts.ScriptTarget.Latest, true);
		sources.set(file, sf);
	}
	return sf;
}

const memberName = (m) =>
	m.name && (ts.isIdentifier(m.name) || ts.isStringLiteral(m.name)) ? m.name.text : null;

/** Members of an interface, followed through `extends` into this file and the shared fit options. */
function interfaceMembers(file, name, seen = new Set()) {
	if (seen.has(name)) return [];
	seen.add(name);
	let found = null;
	for (const candidate of [file, FIT_OPTIONS]) {
		ts.forEachChild(parse(candidate), (n) => {
			if (!found && ts.isInterfaceDeclaration(n) && n.name.text === name) {
				found = n;
			}
		});
		if (found) break;
	}
	if (!found) throw new Error(`interface ${name} not found (looked in ${file})`);
	const out = [];
	for (const m of found.members) {
		const key = memberName(m);
		if (key) out.push(key);
	}
	for (const clause of found.heritageClauses ?? []) {
		for (const parent of clause.types) {
			out.push(...interfaceMembers(file, parent.expression.getText(), seen));
		}
	}
	if (out.length === 0) throw new Error(`interface ${name} has no members`);
	return out;
}

/** String literals in the first parameter of each call signature of an emits interface. */
function emitNames(file, name) {
	const out = [];
	ts.forEachChild(parse(file), (n) => {
		if (!ts.isInterfaceDeclaration(n) || n.name.text !== name) return;
		for (const m of n.members) {
			if (!ts.isCallSignatureDeclaration(m)) continue;
			const type = m.parameters[0]?.type;
			if (!type) continue;
			const literals = ts.isUnionTypeNode(type) ? type.types : [type];
			for (const literal of literals) {
				if (ts.isLiteralTypeNode(literal) && ts.isStringLiteral(literal.literal)) {
					out.push(literal.literal.text);
				}
			}
		}
	});
	if (out.length === 0) throw new Error(`emits interface ${name} not found or empty`);
	return out;
}

/** `readonly x = input(...)` (an `alias` option wins) and `readonly y = output<...>()` members. */
function angularMembers(file, className) {
	const inputs = [];
	const outputs = [];
	ts.forEachChild(parse(file), (n) => {
		if (!ts.isClassDeclaration(n) || n.name?.text !== className) return;
		for (const m of n.members) {
			if (!ts.isPropertyDeclaration(m) || !m.initializer || !ts.isCallExpression(m.initializer)) {
				continue;
			}
			const key = memberName(m);
			const fn = m.initializer.expression.getText();
			if (fn === 'input' || fn === 'input.required') {
				let alias = key;
				for (const arg of m.initializer.arguments) {
					if (!ts.isObjectLiteralExpression(arg)) continue;
					for (const p of arg.properties) {
						if (
							ts.isPropertyAssignment(p) &&
							memberName(p) === 'alias' &&
							ts.isStringLiteral(p.initializer)
						) {
							alias = p.initializer.text;
						}
					}
				}
				inputs.push(alias);
			} else if (fn === 'output') {
				outputs.push(key);
			}
		}
	});
	if (inputs.length === 0) throw new Error(`class ${className} declares no inputs`);
	return { inputs, outputs };
}

const callback = (key) => /^on[A-Z]/.test(key);
const anyCallback = (key) => /^on[A-Za-z]/.test(key);
const split = (keys, isEvent) => ({
	props: keys.filter((key) => !isEvent(key)),
	events: keys.filter(isEvent),
});

/** @returns {Record<'react'|'vue'|'angular'|'svelte'|'vanilla', {props: string[], events: string[]}>} */
export function extractBindings() {
	const pkg = (rel) => join(root, 'packages', rel);
	const vueFile = pkg('vue/src/viewer/types.ts');
	const vue = split(interfaceMembers(vueFile, 'PowerPointViewerProps'), callback);
	const angular = angularMembers(
		pkg('angular/src/viewer/power-point-viewer.component.ts'),
		'PowerPointViewerComponent',
	);
	const angularInputs = split(angular.inputs, callback);
	return {
		react: split(
			interfaceMembers(pkg('react/src/viewer/types-ui.ts'), 'PowerPointViewerProps'),
			callback,
		),
		vue: {
			props: vue.props,
			events: [...vue.events, ...emitNames(vueFile, 'PowerPointViewerEmits')],
		},
		angular: {
			props: angularInputs.props,
			events: [...angularInputs.events, ...angular.outputs],
		},
		svelte: split(
			interfaceMembers(pkg('svelte/src/viewer/types.ts'), 'PowerPointViewerProps'),
			anyCallback,
		),
		vanilla: split(
			interfaceMembers(pkg('vanilla/src/viewer/types.ts'), 'PptxViewerOptions'),
			callback,
		),
	};
}

if (process.argv[1] === fileURLToPath(import.meta.url)) {
	console.log(JSON.stringify(extractBindings(), null, 1));
}
