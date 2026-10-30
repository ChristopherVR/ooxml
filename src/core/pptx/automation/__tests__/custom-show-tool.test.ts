import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';

import { expect, it } from 'vitest';

import { createFileOperations, runAutomationFileTool } from '../../../automation/node';
import { pptxTools } from '../catalogue/index';
import { createTestPptxBytes } from './helpers/create-test-pptx';

const manage = pptxTools.find((t) => t.name === 'manage_custom_shows')!;

interface Shown {
	show?: { id: string; name: string; slideIndices: number[] };
	shows?: { id: string; name: string; slideIndices: number[] }[];
}

it('creates, edits, copies and deletes custom shows and persists them on save', async () => {
	const root = await mkdtemp(join(tmpdir(), 'pptx-custshow-'));
	try {
		await writeFile(join(root, 'deck.pptx'), await createTestPptxBytes(3));
		const files = createFileOperations(root);
		const run = (params: Record<string, unknown>) =>
			runAutomationFileTool(files, manage, { filePath: 'deck.pptx', ...params }) as Promise<Shown>;

		const created = await run({ action: 'create', name: 'Short', slideIndices: [0, 2] });
		expect(created.show).toMatchObject({ name: 'Short', slideIndices: [0, 2] });
		const id = created.show!.id;
		await run({ action: 'addSlides', showId: id, slideIndices: [1] });
		await run({ action: 'moveSlide', showId: id, from: 2, to: 0 });
		const copied = await run({ action: 'copy', showId: id });
		expect(copied.show?.name).toBe('Copy of Short');

		const listed = await run({ action: 'list' });
		expect(listed.shows?.map((s) => s.name)).toEqual(['Short', 'Copy of Short']);
		expect(listed.shows?.[0]?.slideIndices).toEqual([1, 0, 2]);

		await run({ action: 'delete', showId: copied.show!.id });
		expect((await run({ action: 'list' })).shows).toHaveLength(1);
		await expect(run({ action: 'rename', showId: '99', name: 'x' })).rejects.toThrow(
			'No custom show',
		);
		await expect(run({ action: 'create', name: 'Bad', slideIndices: [9] })).rejects.toThrow(
			'out of range',
		);
	} finally {
		await rm(root, { recursive: true, force: true });
	}
});
