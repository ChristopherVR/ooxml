import test from 'node:test';
import assert from 'node:assert/strict';
import { offlineReply } from './offline.js';
import { ACTIONS_FENCE } from './protocol.js';

/** The changes the offline preview proposes for `files` when asked to title-case them. */
async function proposed(files) {
	const full = await offlineReply({
		messages: [{ role: 'user', content: 'User request: Title-case these files' }],
		context: { files },
		signal: new AbortController().signal,
		onText: () => {},
		animate: false,
	});
	const start = full.indexOf('```' + ACTIONS_FENCE);
	if (start < 0) return [];
	const body = full.slice(full.indexOf('\n', start) + 1, full.lastIndexOf('```'));
	return JSON.parse(body).changes;
}

test('the offline preview title-cases text in every product and leaves formulas alone', async () => {
	const changes = await proposed([
		{
			id: 'book',
			kind: 'xlsx',
			content: {
				sheets: [
					{
						index: 0,
						name: 'Budget',
						cells: [
							{ address: 'A1', value: 'launch budget' },
							{ address: 'A2', value: 42 },
							{ address: 'A3', value: 'total cost', formula: 'LOWER("TOTAL COST")' },
							{ address: 'A4', value: 'Already Done' },
						],
					},
				],
			},
		},
		{
			id: 'drawing',
			kind: 'vsdx',
			content: {
				pages: [
					{
						id: '0',
						shapes: [
							{ id: '1', text: 'start here', children: [{ id: '2', text: 'inner step' }] },
							{ id: '3', text: '', children: [] },
						],
					},
				],
			},
		},
	]);
	assert.deepEqual(changes, [
		{
			documentId: 'book',
			type: 'cells',
			sheetIndex: 0,
			cells: [{ address: 'A1', input: 'Launch Budget' }],
		},
		{ documentId: 'drawing', type: 'shape_text', pageId: '0', shapeId: '1', text: 'Start Here' },
		{ documentId: 'drawing', type: 'shape_text', pageId: '0', shapeId: '2', text: 'Inner Step' },
	]);
});
