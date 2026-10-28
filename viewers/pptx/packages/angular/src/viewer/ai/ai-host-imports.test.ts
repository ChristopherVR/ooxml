import { resolve } from 'node:path';

import { describe, expect, it } from 'vitest';

import { walkModuleGraph } from '../../../../../../../scripts/module-graph.mjs';

const forbidden =
	/^(?:ooxml-ui\/pptx\/ai|ai|@ai-sdk\/.+|zod(?:\/.+)?|ooxml-core\/pptx\/automation(?:\/.+)?)$/;

describe('angular assistant runtime boundary', () => {
	// `../ai/public-api.ts` is the `pptx-angular-viewer/ai` secondary entry point.
	for (const entry of [
		'public-api.ts',
		'internals.ts',
		'internal/shared-ai.ts',
		'../ai/public-api.ts',
	]) {
		it(`${entry} keeps the runtime behind an external dynamic import`, async () => {
			const { bare } = await walkModuleGraph(resolve(import.meta.dirname, '../..', entry), {
				followDynamic: false,
			});
			expect(
				bare.filter(
					({ specifier, kind }) => kind !== 'dynamic-import' && forbidden.test(specifier),
				),
			).toStrictEqual([]);
		});
	}

	it('loads the full runtime only from the chat service', async () => {
		const { bare } = await walkModuleGraph(resolve(import.meta.dirname, '../../public-api.ts'));
		const runtime = bare.filter(({ specifier }) => specifier === 'ooxml-ui/pptx/ai');
		expect(runtime).toStrictEqual([
			{
				specifier: 'ooxml-ui/pptx/ai',
				kind: 'dynamic-import',
				importer: resolve(import.meta.dirname, 'ai-chat.service.ts'),
			},
		]);
	});
});
