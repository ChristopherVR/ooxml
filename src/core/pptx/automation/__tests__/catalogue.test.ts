import { mkdtemp, readFile, rm, stat, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';

import { describe, expect, it } from 'vitest';

import { createFileOperations, runAutomationFileTool } from '../../../automation/node';
import { automationTools, UNTRUSTED_NOTE } from '../../../automation/tools/index';
import { PptxHandler } from '../../index';
import { pptxTools } from '../catalogue/index';
import { createTestPptxBytes } from './helpers/create-test-pptx';

const tool = (name: string) => {
	const found = pptxTools.find((candidate) => candidate.name === name);
	if (!found) throw new Error(`No tool ${name}`);
	return found;
};

describe('pptx tool catalogue', () => {
	it('keeps the established names, unique across every format', () => {
		const names = pptxTools.map((t) => t.name);
		expect(names).toHaveLength(74);
		const everything = [
			...names,
			...Object.values(automationTools).flatMap((l) => l.map((t) => t.name)),
		];
		expect(new Set(everything).size).toBe(everything.length);
		for (const t of pptxTools) {
			expect(t.format).toBe('pptx');
			expect(t.input).not.toHaveProperty('filePath');
			expect(t.fileInput).toHaveProperty('filePath');
			expect(t.extensions).toEqual(t.kind === 'read' ? ['.pptx', '.ppt'] : ['.pptx']);
			if (t.kind === 'read') expect(t.description.endsWith(UNTRUSTED_NOTE)).toBe(true);
		}
		expect(tool('find_text').kind).toBe('read');
		expect(tool('manage_comments').description.endsWith(UNTRUSTED_NOTE)).toBe(true);
	});

	it('reads, edits and leaves an unchanged deck unwritten', async () => {
		const root = await mkdtemp(join(tmpdir(), 'pptx-catalogue-'));
		try {
			await writeFile(join(root, 'deck.pptx'), await createTestPptxBytes(2));
			const files = createFileOperations(root);
			const slide = await runAutomationFileTool(files, tool('get_slide'), {
				filePath: 'deck.pptx',
				slideIndex: 0,
			});
			expect(JSON.stringify(slide)).toContain('Slide 1 Title');

			const before = await stat(join(root, 'deck.pptx'));
			const listed = await runAutomationFileTool(files, tool('manage_sections'), {
				filePath: 'deck.pptx',
				action: 'list',
			});
			expect(listed).toMatchObject({ filePath: 'deck.pptx', saved: false });
			expect((await stat(join(root, 'deck.pptx'))).mtimeMs).toBe(before.mtimeMs);

			const added = await runAutomationFileTool(files, tool('add_slide'), {
				filePath: 'deck.pptx',
				outputPath: 'more.pptx',
			});
			expect(added).toMatchObject({ filePath: 'more.pptx', saved: true });
			const reloaded = await new PptxHandler().load(
				(await readFile(join(root, 'more.pptx'))).buffer as ArrayBuffer,
			);
			expect(reloaded.slides).toHaveLength(3);

			await expect(
				runAutomationFileTool(files, tool('convert_to_markdown'), {
					filePath: 'deck.pptx',
					outputDir: '../outside',
				}),
			).rejects.toThrow('relative path');
			const markdown = await runAutomationFileTool(files, tool('convert_to_markdown'), {
				filePath: 'deck.pptx',
			});
			expect(markdown).toMatchObject({ totalSlides: 2 });
		} finally {
			await rm(root, { recursive: true, force: true });
		}
	});

	it('validates parameters itself', async () => {
		const read = tool('get_slide');
		if (read.kind !== 'read') throw new Error('get_slide is not a read tool');
		await expect(read.run(await createTestPptxBytes(1), { slideIndex: -1 })).rejects.toThrow(
			'Invalid get_slide parameters',
		);
		expect(read.jsonSchema()).toMatchObject({ required: ['slideIndex'] });
	});
});
