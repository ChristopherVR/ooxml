import { readFileSync } from 'node:fs';
import { join } from 'node:path';

import { parse } from 'svelte/compiler';
import { describe, expect, it } from 'vitest';

/**
 * The full `ooxml-ui/pptx/ai` barrel bundles the chat session, the tool registry, `zod`, the
 * automation catalogue and a dynamic `import('ai')`. These modules are part of the viewer's main
 * graph (the AI bridge, the panel controller, the change overlay, the Options AI section and the
 * `createAiAssistant` handle registry), so they import the SDK-free helpers from
 * `ooxml-ui/pptx/ai/host`. The barrel stays behind the
 * lazily loaded `AiChatPanel.svelte`. `scripts/check-binding-packages.mjs` checks the same rule
 * on the packed build.
 */
const EAGER_AI_MODULES = [
	'ai-assistant-registry.ts',
	'ai/ai-bridge.ts',
	'ai/ai-panel-controller.svelte.ts',
	'components/ai/AiChangeOverlay.svelte',
	'components/ai/SettingsAiSection.svelte',
];

const IMPORT_SPECIFIER = /\bfrom\s+['"](ooxml-ui\/pptx\/ai(?:\/[^'"]*)?)['"]/gu;

describe('the eager Svelte AI modules import ooxml-ui/pptx/ai/host, not the barrel', () => {
	for (const file of EAGER_AI_MODULES) {
		it(file, () => {
			const source = readFileSync(join(import.meta.dirname, '..', file), 'utf-8');
			const specifiers = [...source.matchAll(IMPORT_SPECIFIER)].map((match) => match[1]);
			expect(specifiers.length).toBeGreaterThan(0);
			expect(specifiers.filter((specifier) => specifier !== 'ooxml-ui/pptx/ai/host')).toStrictEqual(
				[],
			);
		});
	}
});

/**
 * With the barrel out of the main graph, the panel is the first module that reaches it (and
 * `zod`). Vite's dependency scan reads only script blocks, so a template-level `{#await
 * import()}` would hide those dependencies until the panel opens; the dev server then
 * re-optimizes and reloads the page, which drops the click that opened the panel.
 */
describe('AiDock loads the panel from its script block', () => {
	const source = readFileSync(join(import.meta.dirname, '../components/ai/AiDock.svelte'), 'utf-8');
	// Split with the compiler rather than a pattern, so the ranges match what Svelte compiles.
	const ast = parse(source, { modern: true });
	const blocks = [ast.module, ast.instance]
		// An absent block is `undefined` at runtime even though the type says `null`.
		.filter((block): block is NonNullable<typeof block> => block !== null && block !== undefined)
		.toSorted((left, right) => left.start - right.start);
	const script = blocks
		.map((block) => source.slice(block.content.start, block.content.end))
		.join('\n');
	const markup = blocks.reduceRight(
		(rest, block) => rest.slice(0, block.start) + rest.slice(block.end),
		source,
	);

	it('imports AiChatPanel.svelte in the script block', () => {
		expect(script).toMatch(/\bimport\(\s*['"]\.\/AiChatPanel\.svelte['"]\s*\)/u);
	});

	it('has no import() in the template', () => {
		expect(markup).not.toMatch(/\bimport\(/u);
	});

	it('uses the loader of a createAiAssistant handle when it has one', () => {
		expect(script).toMatch(/loadPanel \? loadPanel\(\) : import\(/u);
	});
});
