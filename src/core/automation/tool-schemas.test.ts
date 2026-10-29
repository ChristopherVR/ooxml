import { expect, test } from 'vitest';

import { automationTools } from './tools/index';

// The MCP tool contract for Word, Excel and Visio (PowerPoint's is beside its catalogue in
// pptx/automation/__tests__/tool-schemas.test.ts): a change to a name, a kind, an input or the
// schema version shows up as a diff in the committed snapshot. Raise the tool's `version` when the
// change is incompatible (docs/stability.md), then refresh the snapshot with `vitest -u`.
test('every tool declares an integer schema version of at least 1', () => {
	for (const tool of Object.values(automationTools).flat()) {
		expect(Number.isInteger(tool.version), tool.name).toBe(true);
		expect(tool.version, tool.name).toBeGreaterThanOrEqual(1);
	}
});

for (const format of ['docx', 'xlsx', 'visio'] as const) {
	test(`${format} tool names and input schemas are stable`, async () => {
		const tools = automationTools[format]
			.map((tool) => ({
				name: tool.name,
				kind: tool.kind,
				version: tool.version,
				extensions: tool.extensions,
				inputSchema: tool.jsonSchema(),
			}))
			.sort((a, b) => (a.name < b.name ? -1 : 1));
		await expect(`${JSON.stringify(tools, null, '\t')}\n`).toMatchFileSnapshot(
			`./__snapshots__/tool-schemas.${format}.json`,
		);
	});
}
