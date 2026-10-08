import { expect, it } from 'vitest';
import { readFile, writeFile } from 'node:fs/promises';
import { join } from 'node:path';
import { editVsdx, type VisioTextRangesEdit } from './edit';
import { parseVsdx } from './parser';
import { VisioPackage } from './package';
import type { VisioText } from './model';

const directory = process.env['VISIO_NATIVE_TEXT_RANGES_DIR'];
const characters = (text: VisioText) =>
	text.runs.flatMap((run) =>
		Array.from({ length: run.text.length }, (_, index) => ({ ...run, text: run.text[index] })),
	);
it.skipIf(!directory)(
	'matches native mixed-run character range edits and preserves all source rows across three scales',
	async () => {
		const bytes = await readFile(join(directory!, 'original.vsdx')),
			evidence = JSON.parse(
				(await readFile(join(directory!, 'evidence.json'), 'utf8')).replace(/^\ufeff/, ''),
			) as { cases: { command: VisioTextRangesEdit }[] };
		const result = await editVsdx(
			bytes,
			evidence.cases.map((item) => item.command),
		);
		const actual = await parseVsdx(result.bytes),
			native = await parseVsdx(await readFile(join(directory!, 'native.vsdx')));
		expect(actual.pages).toHaveLength(evidence.cases.length);
		const before = await VisioPackage.open(bytes),
			after = await VisioPackage.open(result.bytes);
		for (let index = 0; index < actual.pages.length; index++) {
			const shape = actual.pages[index]!.shapes[0]!,
				reference = native.pages[index]!.shapes[0]!;
			expect(shape.text.plainText).toBe(reference.text.plainText);
			expect(characters(shape.text), `page${index} characters`).toEqual(characters(reference.text));
			expect(shape.text.paragraphs, `page${index} paragraphs`).toEqual(reference.text.paragraphs);
			expect(shape.geometry).toEqual(reference.geometry);
			expect(shape.transform).toEqual(reference.transform);
			expect(shape.style).toEqual(reference.style);
			const path = `visio/pages/page${index + 1}.xml`;
			const sections = (root: Element) =>
				Array.from(root.getElementsByTagName('Section'))
					.filter((node) => ['Character', 'Paragraph'].includes(node.getAttribute('N')!))
					.map((node) => node.toString());
			expect(sections(await after.readXml(path))).toEqual(sections(await before.readXml(path)));
		}
		for (const path of before.paths())
			if (!result.changedParts.includes(path))
				expect(await after.readBytes(path)).toEqual(await before.readBytes(path));
		await writeFile(join(directory!, 'core.vsdx'), result.bytes);
		await writeFile(
			join(directory!, 'core-reference.json'),
			JSON.stringify({
				cases: evidence.cases.map((item, index) => {
					const text = native.pages[index]!.shapes[0]!.text;
					return {
						command: item.command,
						native: {
							text: text.plainText,
							characters: characters(text).map((run, offset) => ({
								text: run.text,
								style: Number(run.bold) + 2 * Number(run.italic) + 4 * Number(run.underline),
								size: run.fontSize,
								color: run.color,
								alignment: ['left', 'center', 'right', 'justify', 'distributed'].indexOf(
									text.paragraphs!.find(
										(paragraph) => paragraph.start <= offset && paragraph.end >= offset,
									)!.horizontalAlign,
								),
							})),
						},
					};
				}),
			}),
		);
	},
	30000,
);
