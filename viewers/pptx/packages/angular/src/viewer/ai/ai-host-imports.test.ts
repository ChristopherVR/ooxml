import { resolve } from 'node:path';

import { beforeAll, describe, expect, it } from 'vitest';

import { walkModuleGraph, walkModuleGraphs } from '../../../../../../../scripts/module-graph.mjs';

const forbidden =
	/^(?:ooxml-ui\/pptx\/ai|ai|@ai-sdk\/.+|zod(?:\/.+)?|ooxml-core\/pptx\/automation(?:\/.+)?)$/;

type Graph = Awaited<ReturnType<typeof walkModuleGraph>>;

// `../ai/public-api.ts` is the `pptx-angular-viewer/ai` secondary entry point.
const ENTRIES = ['public-api.ts', 'internals.ts', 'internal/shared-ai.ts', '../ai/public-api.ts'];
const entryPath = (entry: string) => resolve(import.meta.dirname, '../..', entry);

describe('angular assistant runtime boundary', () => {
	let initial: Graph[] = [];
	let full: Graph | undefined;

	// Each walk bundles the package's sources with esbuild. Two passes cover every case below (the
	// entries share most of their graph); on a loaded runner they take well past one test's budget.
	beforeAll(async () => {
		[initial, full] = await Promise.all([
			walkModuleGraphs(ENTRIES.map(entryPath), { followDynamic: false }),
			walkModuleGraph(entryPath('public-api.ts')),
		]);
	}, 120_000);

	ENTRIES.forEach((entry, index) => {
		it(`${entry} keeps the runtime behind an external dynamic import`, () => {
			expect(
				initial[index]!.bare.filter(
					({ specifier, kind }) => kind !== 'dynamic-import' && forbidden.test(specifier),
				),
			).toStrictEqual([]);
		});
	});

	it('loads the full runtime only from the chat service', () => {
		const runtime = full!.bare.filter(({ specifier }) => specifier === 'ooxml-ui/pptx/ai');
		expect(runtime).toStrictEqual([
			{
				specifier: 'ooxml-ui/pptx/ai',
				kind: 'dynamic-import',
				importer: resolve(import.meta.dirname, 'ai-chat.service.ts'),
			},
		]);
	});
});
