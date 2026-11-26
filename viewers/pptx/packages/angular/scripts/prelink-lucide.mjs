import { createHash, randomUUID } from 'node:crypto';
import { existsSync, mkdirSync, readFileSync, renameSync, rmSync, writeFileSync } from 'node:fs';
import { createRequire } from 'node:module';
import { dirname, join } from 'node:path';

import ts from 'typescript';

const lucide = '@lucide/angular';
const require = createRequire(import.meta.url);

function importedIcons(root) {
	const imports = new Set();
	const files = ts.sys.readDirectory(
		root,
		['.ts', '.tsx', '.mts', '.cts', '.js', '.mjs', '.cjs'],
		undefined,
		['src/**/*', 'ai/**/*'],
	);
	for (const file of files) {
		const source = readFileSync(file, 'utf8');
		if (!source.includes(lucide)) {
			continue;
		}
		const tree = ts.createSourceFile(file, source, ts.ScriptTarget.Latest, true);
		const needsWholeBarrel = ts.forEachChild(tree, function visit(node) {
			if (!ts.isStringLiteralLike(node) || node.text !== lucide) {
				return ts.forEachChild(node, visit);
			}
			const declaration = node.parent;
			if (!ts.isImportDeclaration(declaration) || declaration.moduleSpecifier !== node) {
				return true;
			}
			const clause = declaration.importClause;
			if (clause?.isTypeOnly) {
				return;
			}
			const bindings = clause?.namedBindings;
			if (clause?.name || !bindings || !ts.isNamedImports(bindings)) {
				return true;
			}
			for (const specifier of bindings.elements) {
				if (!specifier.isTypeOnly) {
					imports.add((specifier.propertyName ?? specifier.name).text);
				}
			}
		});
		if (needsWholeBarrel) {
			return undefined;
		}
	}
	return [...imports].sort();
}

// A partial icon's static definitions and its following metadata call both refer
// to that class. Keep them together so esbuild can remove unreferenced icons and
// retain all declarations and metadata for the icons the binding imports. Leave
// any other statement in the installed entry where Lucide put it.
function markTreeShakeableIcons(filename, source) {
	const tree = ts.createSourceFile(filename, source, ts.ScriptTarget.Latest);
	const parts = [];
	let offset = 0;
	for (let index = 0; index < tree.statements.length - 1; index++) {
		const statement = tree.statements[index];
		const metadata = tree.statements[index + 1];
		if (!ts.isVariableStatement(statement) || !ts.isExpressionStatement(metadata)) {
			continue;
		}
		const { declarations } = statement.declarationList;
		const declaration = declarations[0];
		const icon = declaration.initializer;
		if (declarations.length !== 1 || !ts.isIdentifier(declaration.name) || !icon) {
			continue;
		}
		if (
			!ts.isClassExpression(icon) ||
			icon.heritageClauses?.[0]?.getText(tree) !== 'extends LucideIconBase' ||
			icon.members.some(ts.isClassStaticBlockDeclaration)
		) {
			continue;
		}
		const name = declaration.name.text;
		const call = metadata.expression;
		if (
			!ts.isCallExpression(call) ||
			call.expression.getText(tree) !== 'i0.ɵɵngDeclareClassMetadata'
		) {
			continue;
		}
		const argument = call.arguments[0];
		if (!argument || !ts.isObjectLiteralExpression(argument)) {
			continue;
		}
		const type = argument.properties.find(
			(property) => ts.isPropertyAssignment(property) && property.name.getText(tree) === 'type',
		);
		if (!type || type.initializer.getText(tree) !== name) {
			continue;
		}
		parts.push(
			source.slice(offset, statement.getStart(tree)),
			`var ${name} = /* @__PURE__ */ (() => { ${statement.getText(tree)}\n${metadata.getText(tree)}\nreturn ${name}; })();`,
		);
		offset = metadata.end;
	}
	parts.push(source.slice(offset));
	return parts.join('');
}

async function treeShake(filename, source, icons) {
	if (!icons) {
		return source;
	}
	const { build } = await import('esbuild');
	const result = await build({
		stdin: {
			contents: `export { ${icons.join(', ')} } from ${JSON.stringify(lucide)};`,
			resolveDir: import.meta.dirname,
		},
		bundle: true,
		packages: 'external',
		write: false,
		format: 'esm',
		platform: 'node',
		charset: 'utf8',
		plugins: [
			{
				name: 'lucide-partial-declarations',
				setup(builder) {
					builder.onResolve({ filter: /^@lucide\/angular$/ }, () => ({ path: filename }));
					builder.onLoad({ filter: /./ }, ({ path }) => {
						if (path === filename) {
							return {
								contents: markTreeShakeableIcons(filename, source),
								resolveDir: dirname(filename),
							};
						}
					});
				},
			},
		],
	});
	return result.outputFiles[0].text;
}

export async function prelinkLucide(root) {
	const filename = require.resolve(lucide);
	const icons = importedIcons(root);
	if (icons?.length === 0) {
		return filename;
	}
	const source = readFileSync(filename, 'utf8');
	const versions = [lucide, '@angular/compiler-cli', '@babel/core', 'esbuild', 'typescript'].map(
		(dependency) => require(`${dependency}/package.json`).version,
	);
	const hash = createHash('sha256')
		.update(JSON.stringify([versions, icons]))
		.update(source)
		.update(readFileSync(new URL('./prelink-lucide.mjs', import.meta.url)))
		.digest('hex');

	// Keep the cached ESM module inside the binding's node_modules. Vitest can
	// externalize it, so native Angular and Lucide remain shared across files.
	const cacheDir = join(root, 'node_modules', '.cache', 'lucide-linker');
	const cached = join(cacheDir, `${hash}.mjs`);
	if (existsSync(cached)) {
		return cached;
	}

	const { transformAsync } = await import('@babel/core');
	const { default: linker } = await import('@angular/compiler-cli/linker/babel');
	const result = await transformAsync(await treeShake(filename, source, icons), {
		filename,
		babelrc: false,
		configFile: false,
		inputSourceMap: false,
		plugins: [[linker, { sourceMapping: false }]],
	});
	if (!result?.code) {
		throw new Error(`Angular linker produced no output for ${filename}`);
	}
	mkdirSync(cacheDir, { recursive: true });
	const temporary = `${cached}.${randomUUID()}.tmp`;
	try {
		writeFileSync(temporary, result.code);
		renameSync(temporary, cached);
	} finally {
		rmSync(temporary, { force: true });
	}
	return cached;
}
