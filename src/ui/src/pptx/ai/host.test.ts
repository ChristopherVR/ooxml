import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';

import { describe, expect, it } from 'vitest';

import type { BareImport } from '../../../../../scripts/module-graph.mjs';
import { walkModuleGraph } from '../../../../../scripts/module-graph.mjs';
import * as barrel from './index';
import * as host from './host';

/**
 * `ooxml-ui/pptx/ai/host` is the SDK-free part of the assistant core that the viewers'
 * main entries import. A runtime import of the optional `ai` SDK, of `zod` or of the
 * automation tool catalogue anywhere in its graph would put the assistant runtime back
 * into every host's initial bundle (and make the host's bundler resolve `ai`).
 */
const FORBIDDEN = /^(?:ai|@ai-sdk\/.+|zod(?:\/.+)?|ooxml-core\/pptx\/automation(?:\/.+)?)$/;

const graphOf = (file: string) => walkModuleGraph(fileURLToPath(new URL(file, import.meta.url)));

const forbiddenImports = (bare: readonly BareImport[]) =>
	bare
		.filter(({ specifier }) => FORBIDDEN.test(specifier))
		.map(({ specifier, kind, importer }) => `${kind} of ${specifier} in ${importer}`);

describe('ooxml-ui/pptx/ai/host', () => {
	it('reaches no runtime import of ai, zod or the automation catalogue', async () => {
		const { files, bare } = await graphOf('./host.ts');
		expect(forbiddenImports(bare)).toStrictEqual([]);
		// The walk followed the helpers' own imports, not only host.ts.
		expect(files.length).toBeGreaterThan(5);
	});

	it('never goes through the full ai barrel', async () => {
		const { files } = await graphOf('./host.ts');
		const barrelFile = fileURLToPath(new URL('./index.ts', import.meta.url));
		expect(files).not.toContain(barrelFile);
	});

	it('the same walk finds the assistant runtime behind the full barrel', async () => {
		const found = new Set(
			(await graphOf('./index.ts')).bare
				.filter(({ specifier }) => FORBIDDEN.test(specifier))
				.map(({ specifier, kind }) => `${specifier} ${kind}`),
		);
		expect(found).toContain('ai dynamic-import');
		expect(found).toContain('zod import-statement');
		expect(found).toContain('ooxml-core/pptx/automation import-statement');
	});

	it('exports the same values as the full barrel', () => {
		const names = Object.keys(host).sort();
		expect(names).toStrictEqual(
			[
				'AI_CHANGE_ATTR',
				'AI_TOOL_FLASH_MS',
				'aiChangeAnimationCss',
				'aiFixDirective',
				'applyElementUpdate',
				'applyShapeStyleUpdate',
				'applyTextUpdate',
				'buildChatLogExport',
				'buildChatLogMarkdown',
				'changeGhostStyle',
				'collectStoredChats',
				'computeFocusTargets',
				'createAiChangeAnimator',
				'createChatHistoryStore',
				'createPptxAiBridge',
				'createPptxAiPanelController',
				'deckDataFieldChanged',
				'diffChangedElements',
				'exportAiChatLogs',
				'focusTargetChips',
				'isAiFocusPinned',
				'isTwoTableFocus',
				'mergeTablesDirective',
				'pptxAiDeckTitle',
				'resolveAiEffectiveTargets',
				'resolveChangeAnimationConfig',
				'toLogChat',
				'toolCallLine',
			].sort(),
		);
		for (const name of names) {
			expect(host[name as keyof typeof host], name).toBe(barrel[name as keyof typeof barrel]);
		}
	});

	it('is built as its own entry and published through the ./pptx/* export', () => {
		const tsup = readFileSync(new URL('../../../tsup.pptx.config.ts', import.meta.url), 'utf8');
		expect(tsup).toContain("'ai/host': 'src/pptx/ai/host.ts'");
		const manifest = JSON.parse(
			readFileSync(new URL('../../../package.json', import.meta.url), 'utf8'),
		) as { exports: Record<string, { import: string; require: string; types: string }> };
		expect(manifest.exports['./pptx/ai/host']).toBeUndefined();
		expect(manifest.exports['./pptx/*']).toStrictEqual({
			types: './dist/pptx/*.d.ts',
			import: './dist/pptx/*.mjs',
			require: './dist/pptx/*.cjs',
		});
	});
});
