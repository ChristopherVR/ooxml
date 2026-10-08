import { expect, it } from 'vitest';
import { readFile, writeFile } from 'node:fs/promises';
import { join } from 'node:path';
import { editVsdx } from '../edit';
import { parseVsdx } from '../parser';
import { VisioPackage } from '../package';
import {
	visioTextReplaceCommands,
	visioTextReplacePlan,
	type VisioTextOccurrence,
} from './text-replace';

const directory = process.env['VISIO_NATIVE_TEXT_REPLACE_DIR'];
it.skipIf(!directory)(
	'matches native Characters-range literal replacement across three drawing scales',
	async () => {
		const bytes = await readFile(join(directory!, 'original.vsdx')),
			document = await parseVsdx(bytes);
		const evidence = JSON.parse(
			(await readFile(join(directory!, 'evidence.json'), 'utf8')).replace(/^\ufeff/, ''),
		) as {
			cases: {
				pageId: string;
				shapeId: string;
				request: { query: string; replacement: string; mode: 'all' | 'current' };
				current?: VisioTextOccurrence;
				occurrences: VisioTextOccurrence[];
				nativeText: string;
			}[];
		};
		const edits = evidence.cases.flatMap((item) => {
			const plan = visioTextReplacePlan(document, {
				...item.request,
				matchCase: true,
				scope: 'current-page',
				pageId: item.pageId,
				...(item.current ? { current: item.current } : {}),
			});
			expect(plan.occurrences).toEqual(item.occurrences);
			return visioTextReplaceCommands(document, plan);
		});
		const result = await editVsdx(bytes, edits),
			actual = await parseVsdx(result.bytes),
			native = await parseVsdx(await readFile(join(directory!, 'native.vsdx')));
		for (let index = 0; index < evidence.cases.length; index++) {
			const shape = actual.pages[index]!.shapes[0]!,
				reference = native.pages[index]!.shapes[0]!;
			expect(shape.text.plainText).toBe(evidence.cases[index]!.nativeText);
			expect(shape.text.plainText).toBe(reference.text.plainText);
			expect(shape.geometry).toEqual(reference.geometry);
			expect(shape.transform).toEqual(reference.transform);
			expect(shape.style).toEqual(reference.style);
		}
		const before = await VisioPackage.open(bytes),
			after = await VisioPackage.open(result.bytes);
		for (const path of before.paths())
			if (!result.changedParts.includes(path))
				expect(await after.readBytes(path)).toEqual(await before.readBytes(path));
		await writeFile(join(directory!, 'core.vsdx'), result.bytes);
	},
	30000,
);
