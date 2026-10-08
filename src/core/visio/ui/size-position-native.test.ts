import { readFile, writeFile } from 'node:fs/promises';
import { join } from 'node:path';
import { describe, expect, it } from 'vitest';
import { editVsdx } from '../edit';
import { parseVsdx } from '../parser';
import { visioSizePositionCommand, type VisioSizePositionField } from './size-position';

const directory = process.env.VISIO_NATIVE_SIZE_POSITION_DIR;
describe.skipIf(!directory)('native numeric pin, size and angle edits', () => {
	it('matches native cell edits at three scales and writes core output for independent reopen', async () => {
		const cases = JSON.parse(await readFile(join(directory!, 'cases.json'), 'utf8')) as {
			name: string;
			field: VisioSizePositionField;
			value: number;
		}[];
		expect(cases).toHaveLength(15);
		for (const item of cases) {
			const bytes = await readFile(join(directory!, `${item.name}-original.vsdx`));
			const source = (await parseVsdx(bytes)).pages[0]!;
			const command = visioSizePositionCommand(source, '1', item.field, item.value)!;
			expect(command).toHaveLength(1);
			const saved = await editVsdx(bytes, command);
			const actual = (await parseVsdx(saved.bytes)).pages[0]!.shapes[0]!;
			const expected = (
				await parseVsdx(await readFile(join(directory!, `${item.name}-native.vsdx`)))
			).pages[0]!.shapes[0]!;
			expect(actual.width, item.name).toBeCloseTo(expected.width, 10);
			expect(actual.height, item.name).toBeCloseTo(expected.height, 10);
			for (const [index, value] of actual.transform.entries())
				expect(value, `${item.name}/${index}`).toBeCloseTo(expected.transform[index]!, 10);
			expect(actual.text.plainText, item.name).toBe(expected.text.plainText);
			await writeFile(join(directory!, `${item.name}-core.vsdx`), saved.bytes);
		}
	});
});
