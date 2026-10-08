import { readFile, writeFile } from 'node:fs/promises';
import { existsSync } from 'node:fs';
import { join } from 'node:path';
import { describe, expect, it } from 'vitest';
import { createVsdx } from './create-document';
import { editVsdx } from './edit';
import type { VisioEdit } from './edit-commands';
import { parseVsdx } from './parser';

const directory = process.env.VISIO_NATIVE_TEXT_PARAGRAPHS_DIR;
const reopened = directory && existsSync(join(directory, 'core-resaved.vsdx'));
type Case = { id: string; text: string; mode?: string };
describe.skipIf(!directory)('native logical text and paragraph terminators', () => {
	it('matches native saved logical text, including intentional trailing blank paragraphs', async () => {
		const evidence = JSON.parse(
			(await readFile(join(directory!, 'native-evidence.json'), 'utf8')).trim(),
		) as {
			application: string;
			cases: Case[];
		};
		expect(evidence.application).toBe('Microsoft Visio');
		expect(evidence.cases).toHaveLength(10);
		const native = await parseVsdx(await readFile(join(directory!, 'native.vsdx')));
		for (const item of evidence.cases) {
			const text = native.pages[0]!.shapes.find((shape) => shape.id === item.id)!.text;
			expect(text.plainText, item.id).toBe(item.text);
			expect(text.runs.map((run) => run.text).join('')).toBe(item.text);
			expect(text.paragraphs?.length).toBe(item.text ? item.text.split('\n').length : 0);
		}
	});
	it('writes source edits with exact logical text for owned native reopen', async () => {
		const evidence = JSON.parse(
			(await readFile(join(directory!, 'native-evidence.json'), 'utf8')).trim(),
		) as {
			cases: Case[];
		};
		const cases: Case[] = [];
		const commands: VisioEdit[] = [];
		for (const mode of [
			'create-rectangle',
			'create-ellipse',
			'create-text-box',
			'replace-plain-text',
		] as const)
			for (const item of evidence.cases) {
				const id = String(cases.length + 1);
				cases.push({ id, text: item.text, mode });
				commands.push({
					type: mode === 'replace-plain-text' ? 'create-rectangle' : mode,
					pageId: '0',
					shapeId: id,
					x: 2,
					y: 3,
					width: 2,
					height: 1,
					text: mode === 'replace-plain-text' ? 'Before replacement' : item.text,
				});
				if (mode === 'replace-plain-text')
					commands.push({ type: mode, pageId: '0', shapeId: id, text: item.text });
			}
		const saved = await editVsdx(await createVsdx(), commands);
		const model = await parseVsdx(saved.bytes);
		for (const item of cases)
			expect(model.pages[0]!.shapes.find((shape) => shape.id === item.id)!.text.plainText).toBe(
				item.text,
			);
		await writeFile(join(directory!, 'core.vsdx'), saved.bytes);
		await writeFile(join(directory!, 'core-cases.json'), JSON.stringify(cases));
	});
	it.skipIf(!reopened)(
		'retains all logical text after native open and save of core outputs',
		async () => {
			const evidence = JSON.parse(
				(await readFile(join(directory!, 'reopen-evidence.json'), 'utf8')).trim(),
			) as {
				cases: Case[];
			};
			expect(evidence.cases).toHaveLength(40);
			const native = await parseVsdx(await readFile(join(directory!, 'core-resaved.vsdx')));
			for (const item of evidence.cases)
				expect(
					native.pages[0]!.shapes.find((shape) => shape.id === item.id)!.text.plainText,
					item.mode,
				).toBe(item.text);
		},
	);
});
