import { readFileSync } from 'node:fs';
import { join, relative, sep } from 'node:path';

import { describe, expect, it } from 'vitest';

import { walkModuleGraph } from '../../../../../../../scripts/module-graph.mjs';

/**
 * The full `ooxml-ui/pptx/ai` barrel bundles the chat session, the tool registry, `zod`, the
 * automation catalogue and a dynamic `import('ai')`. The main entries use only the SDK-free
 * helpers, from `ooxml-ui/pptx/ai/host`; the barrel belongs to the lazily loaded panel chunk
 * (`ai-panel` and what it imports), so a host that never opens the assistant never loads or
 * bundles it eagerly.
 */
describe('the vanilla entries load AI helpers from ooxml-ui/pptx/ai/host only', () => {
	const packageRoot = join(import.meta.dirname, '../../..');
	const staticGraph = (file: string) =>
		walkModuleGraph(join(packageRoot, file), { followDynamic: false });
	// Forward slashes on every platform, to match the expectations below.
	const relativeName = (file: string) => relative(packageRoot, file).split(sep).join('/');
	const ENTRIES = ['src/index.ts', 'src/internals.ts', 'src/ai.ts'];

	it('lists every built entry', () => {
		const tsup = readFileSync(join(packageRoot, 'tsup.config.ts'), 'utf-8');
		const entries = /entry:\s*\[([^\]]*)\]/u.exec(tsup)?.[1] ?? '';
		expect([...entries.matchAll(/'([^']+)'/gu)].map((match) => match[1])).toStrictEqual(ENTRIES);
	});

	for (const entry of ENTRIES) {
		it(`${entry} has no runtime import of the ai barrel or the ai SDK outside the panel chunk`, async () => {
			const { bare } = await staticGraph(entry);
			const offenders = bare
				.filter(({ specifier }) => specifier === 'ai' || specifier === 'ooxml-ui/pptx/ai')
				.map(
					({ specifier, kind, importer }) => `${kind} of ${specifier} in ${relativeName(importer)}`,
				);
			expect(offenders).toStrictEqual([]);
		});
	}

	it('reaches the AI panel only through its lazy boundary', async () => {
		const { files, lazy, bare } = await staticGraph('src/index.ts');
		const panel = join(import.meta.dirname, 'ai-panel.ts');
		expect(files).not.toContain(panel);
		expect(lazy).toContain(panel);
		// The bridge, the panel controller, the overlay, the Options section and the
		// `createAiAssistant` handle registry stay eager, on the SDK-free leaf.
		const hostImporters = [
			...new Set(
				bare
					.filter(({ specifier }) => specifier === 'ooxml-ui/pptx/ai/host')
					.map(({ importer }) => relativeName(importer)),
			),
		].sort();
		expect(hostImporters).toStrictEqual([
			'src/viewer/ai-assistant-registry.ts',
			'src/viewer/ai/ai-bridge.ts',
			'src/viewer/ai/ai-change-overlay.ts',
			'src/viewer/ai/ai-panel-controller.ts',
			'src/viewer/ai/ai-settings-section.ts',
		]);
		// The panel chunk is where the full barrel is used, so the walk does see it there.
		const panelImports = (await staticGraph('src/viewer/ai/ai-panel.ts')).bare.map(
			({ specifier }) => specifier,
		);
		expect(panelImports).toContain('ooxml-ui/pptx/ai');
	});

	it('pptx-vanilla-viewer/ai reaches the same panel only through its lazy boundary', async () => {
		const { files, lazy } = await staticGraph('src/ai.ts');
		const panel = join(import.meta.dirname, 'ai-panel.ts');
		expect(files).not.toContain(panel);
		expect(lazy).toStrictEqual([panel]);
	});
});
