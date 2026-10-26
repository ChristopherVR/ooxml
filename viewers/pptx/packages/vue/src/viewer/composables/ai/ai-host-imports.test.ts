import { readFileSync } from 'node:fs';
import { join } from 'node:path';

import { describe, expect, it } from 'vitest';

/**
 * The full `ooxml-ui/pptx/ai` barrel bundles the chat session, the tool registry, `zod`, the
 * automation catalogue and a dynamic `import('ai')`. These modules are part of the viewer's main
 * graph (the AI bridge, the panel controller, the change overlay and the Options AI tab), so
 * they import the SDK-free helpers from `ooxml-ui/pptx/ai/host`. The barrel stays behind the
 * lazily loaded `AiChatPanel.vue`. `scripts/check-binding-packages.mjs` checks the same rule on
 * the packed build.
 */
const EAGER_AI_MODULES = [
	'composables/ai/useAiBridge.ts',
	'composables/ai/useAiPanelController.ts',
	'composables/ai/ai-log-export.ts',
	'components/SettingsAiTab.vue',
	'components/ai/AiChangeOverlay.vue',
];

const IMPORT_SPECIFIER = /\bfrom\s+['"](ooxml-ui\/pptx\/ai(?:\/[^'"]*)?)['"]/gu;

describe('the eager Vue AI modules import ooxml-ui/pptx/ai/host, not the barrel', () => {
	for (const file of EAGER_AI_MODULES) {
		it(file, () => {
			const source = readFileSync(join(import.meta.dirname, '../..', file), 'utf-8');
			const specifiers = [...source.matchAll(IMPORT_SPECIFIER)].map((match) => match[1]);
			expect(specifiers.length).toBeGreaterThan(0);
			expect(specifiers.filter((specifier) => specifier !== 'ooxml-ui/pptx/ai/host')).toStrictEqual(
				[],
			);
		});
	}
});
