import { expect, it } from 'vitest';

import { pptxTools } from '../catalogue/index';

// The PowerPoint half of the MCP tool contract; the other formats are in
// src/core/automation/tool-schemas.test.ts. Raise a tool's `version` when a change is
// incompatible (docs/stability.md), then refresh the snapshot with `vitest -u`.
it('pptx tool names and input schemas are stable', async () => {
	for (const tool of pptxTools) {
		expect(Number.isInteger(tool.version), tool.name).toBe(true);
		expect(tool.version, tool.name).toBeGreaterThanOrEqual(1);
	}
	const tools = pptxTools
		.map((tool) => ({
			name: tool.name,
			kind: tool.kind,
			version: tool.version,
			extensions: tool.extensions,
			inputSchema: tool.jsonSchema(),
		}))
		.sort((a, b) => (a.name < b.name ? -1 : 1));
	await expect(`${JSON.stringify(tools, null, '\t')}\n`).toMatchFileSnapshot(
		'./tool-schemas.pptx.json',
	);
});
