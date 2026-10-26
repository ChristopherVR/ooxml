import { readFileSync } from 'node:fs';
import { join } from 'node:path';

import * as ts from '@typescript/typescript6';
import { rollup } from 'rollup';
import { describe, expect, it } from 'vitest';

import { walkModuleGraph } from '../../../../../../../../scripts/module-graph.mjs';

/**
 * Regression for issue #143 (and its sibling in `pptx-viewer`): `@ai-sdk/react`
 * is an optional peer (`peerDependenciesMeta['@ai-sdk/react'].optional ===
 * true`), so a consumer who has not installed it gets an empty stub module
 * from their bundler's optional-peer-dependency handling. A STATIC named
 * import of `useChat` from that module (`import { useChat } from
 * '@ai-sdk/react'`) asks Rollup to validate the binding at link time, which
 * fails the CONSUMER's own production build with `"useChat" is not exported`,
 * even though the AI panel is only ever reached once the user opens it.
 *
 * `AiChatPanelLazy.tsx` resolves `useChat` with a dynamic `import()` inside
 * its `React.lazy` factory and threads it down as a prop; `useAiConversation`
 * receives it as a parameter instead of importing the SDK itself. This test
 * reads both files' source directly rather than the built bundle: it is the
 * static `import` declaration that trips Rollup's link-time validation, so
 * asserting on it here is a faithful (and far cheaper) proxy for building the
 * whole package and grepping its dist chunk.
 */
describe('the @ai-sdk/react peer is never statically imported', () => {
	const lazySource = readFileSync(join(import.meta.dirname, 'AiChatPanelLazy.tsx'), 'utf-8');
	const hookSource = readFileSync(
		join(import.meta.dirname, '../../hooks/ai/useAiConversation.ts'),
		'utf-8',
	);

	it('loads useChat through a dynamic import() in AiChatPanelLazy.tsx', () => {
		expect(lazySource).toMatch(/import\(\s*['"]@ai-sdk\/react['"]\s*\)/u);
	});

	it('has no static value import of @ai-sdk/react anywhere in the AI chat panel chunk', () => {
		for (const [label, source] of [
			['AiChatPanelLazy.tsx', lazySource],
			['useAiConversation.ts', hookSource],
		] as const) {
			const staticValueImports = [
				...source.matchAll(/^import\s+(?!type\s).*from\s+['"]@ai-sdk\/react['"]/gmu),
			];
			expect(
				staticValueImports,
				`${label} should not statically import @ai-sdk/react`,
			).toStrictEqual([]);
		}
	});

	it('does not statically import runtime values from the optional ai peer', () => {
		const messagePartsSource = readFileSync(
			join(import.meta.dirname, 'ai-message-parts.ts'),
			'utf-8',
		);
		const staticAiImports = [
			...messagePartsSource.matchAll(/^import\s+(?!type\s).*from\s+['"]ai['"]/gmu),
		];
		expect(
			staticAiImports,
			'AI message helpers must use the shared runtime-free implementation',
		).toStrictEqual([]);
	});

	it('bundles the React and Vue message helpers with an empty optional-ai stub', async () => {
		const sharedUiParts = join(
			import.meta.dirname,
			'../../../../../../../../src/ui/src/pptx/ai/ui-parts.ts',
		);
		const plugin = {
			name: 'optional-ai-peer-fixture',
			resolveId(source: string) {
				if (source === 'ooxml-ui/pptx/ai') {
					return sharedUiParts;
				}
				if (source === 'ai') {
					return '\0empty-ai-peer';
				}
				return null;
			},
			load(id: string) {
				if (id === '\0empty-ai-peer') {
					return 'export {};';
				}
				return null;
			},
			async transform(code: string, id: string) {
				if (id.endsWith('.ts')) {
					return ts.transpileModule(code, {
						compilerOptions: { module: ts.ModuleKind.ESNext, target: ts.ScriptTarget.ES2022 },
					}).outputText;
				}
				return null;
			},
		};
		const entries = [
			join(import.meta.dirname, 'ai-message-parts.ts'),
			join(import.meta.dirname, '../../../../../vue/src/viewer/composables/ai/message-parts.ts'),
		];
		for (const entry of entries) {
			const bundle = await rollup({ input: entry, plugins: [plugin] });
			try {
				await bundle.generate({ format: 'es' });
			} finally {
				await bundle.close();
			}
		}
	});

	it('confirms the empty ai stub rejects a static named import', async () => {
		const plugin = {
			name: 'optional-ai-peer-negative-control',
			resolveId(source: string) {
				if (source === 'virtual:consumer') {
					return '\0consumer';
				}
				if (source === 'ai') {
					return '\0empty-ai-peer';
				}
				return null;
			},
			load(id: string) {
				if (id === '\0consumer') {
					return "import { isToolUIPart } from 'ai'; console.log(isToolUIPart);";
				}
				if (id === '\0empty-ai-peer') {
					return 'export {};';
				}
				return null;
			},
		};
		await expect(async () => {
			const bundle = await rollup({ input: 'virtual:consumer', plugins: [plugin] });
			try {
				await bundle.generate({ format: 'es' });
			} finally {
				await bundle.close();
			}
		}).rejects.toThrow(/isToolUIPart.*not exported/iu);
	});
});

/**
 * The full `ooxml-ui/pptx/ai` barrel bundles the chat session, the tool registry, `zod`, the
 * automation catalogue and a dynamic `import('ai')`. The main entries use only the SDK-free
 * helpers, from `ooxml-ui/pptx/ai/host`; the barrel belongs to the lazily loaded panel chunk
 * (`AiChatPanel` and what it imports), so a host that never opens the assistant never loads
 * or bundles it eagerly.
 */
describe('the main entries load AI helpers from ooxml-ui/pptx/ai/host only', () => {
	const packageRoot = join(import.meta.dirname, '../../../..');
	const staticGraph = (file: string) =>
		walkModuleGraph(join(packageRoot, file), { followDynamic: false });
	const relativeName = (file: string) => file.slice(packageRoot.length + 1);
	const ENTRIES = ['src/index.ts', 'src/viewer/index.ts', 'src/i18n.ts', 'src/internals.ts'];

	it('lists every built entry', () => {
		const tsup = readFileSync(join(packageRoot, 'tsup.config.ts'), 'utf-8');
		const entries = /entry:\s*\[([^\]]*)\]/u.exec(tsup)?.[1] ?? '';
		expect([...entries.matchAll(/'([^']+)'/gu)].map((match) => match[1])).toStrictEqual(ENTRIES);
	});

	for (const entry of ENTRIES) {
		it(`${entry} has no runtime import of the ai barrel or the ai SDK outside the panel chunk`, async () => {
			const { bare } = await staticGraph(entry);
			const offenders = bare
				.filter(({ specifier, kind }) =>
					specifier === '@ai-sdk/react'
						? kind !== 'dynamic-import'
						: specifier === 'ai' || specifier === 'ooxml-ui/pptx/ai',
				)
				.map(
					({ specifier, kind, importer }) => `${kind} of ${specifier} in ${relativeName(importer)}`,
				);
			expect(offenders).toStrictEqual([]);
		});
	}

	it('reaches the AI panel only through its lazy boundary', async () => {
		const { files, lazy, bare } = await staticGraph('src/index.ts');
		const panel = join(import.meta.dirname, 'AiChatPanel.tsx');
		expect(files).not.toContain(panel);
		expect(lazy).toContain(panel);
		// The overlays, the bridge and the Options tab stay eager, on the SDK-free leaf.
		const hostImporters = bare
			.filter(({ specifier }) => specifier === 'ooxml-ui/pptx/ai/host')
			.map(({ importer }) => relativeName(importer))
			.sort();
		expect(hostImporters).toStrictEqual([
			'src/viewer/components/SettingsAiTab.tsx',
			'src/viewer/components/ai/AiChangeOverlay.tsx',
			'src/viewer/hooks/ai/useAiBridge.ts',
			'src/viewer/hooks/ai/useAiPanelController.ts',
			'src/viewer/utils/ai-log-export.ts',
		]);
		// The panel chunk is where the full barrel is used, so the walk does see it there.
		const panelImports = (await staticGraph('src/viewer/components/ai/AiChatPanel.tsx')).bare.map(
			({ specifier }) => specifier,
		);
		expect(panelImports).toContain('ooxml-ui/pptx/ai');
	});
});
