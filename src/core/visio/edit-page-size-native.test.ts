import { expect, it } from 'vitest';
import { readFile, writeFile, mkdir } from 'node:fs/promises';
import { basename, join } from 'node:path';
import { editVsdx } from './edit';
import { parseVsdx } from './parser';
import { VisioPackage } from './package';

interface NativeCase {
	scale: number;
	operation: string;
	beforeFile: string;
	afterFile: string;
	after: { foreground: { id: string; cells: Record<string, { value: number }> } };
}
for (const variable of [
	'VISIO_NATIVE_PAGE_SIZE_DIR',
	'VISIO_NATIVE_PAGE_SIZE_DEPENDENT_DIR',
	'VISIO_NATIVE_PAGE_SIZE_MODES_DIR',
	'VISIO_NATIVE_PAGE_SIZE_FACTORY_DIR',
]) {
	const directory = process.env[variable];
	it.skipIf(!directory)(
		`matches native fixed page dimensions or refuses dependent caches (${variable})`,
		async () => {
			const evidence = JSON.parse(
				(await readFile(join(directory!, 'evidence.json'), 'utf8')).replace(/^\ufeff/, ''),
			) as { cases: NativeCase[] };
			const output = join(directory!, 'core');
			await mkdir(output, { recursive: true });
			let compared = 0;
			for (const item of evidence.cases) {
				if (!item.operation.startsWith('drawing-')) continue;
				const original = await readFile(item.beforeFile);
				const command = {
					type: 'set-page-size' as const,
					pageId: String(item.after.foreground.id),
					width: item.after.foreground.cells.PageWidth!.value / item.scale,
					height: item.after.foreground.cells.PageHeight!.value / item.scale,
				};
				if (variable.endsWith('DEPENDENT_DIR')) {
					await expect(editVsdx(original, [command])).rejects.toMatchObject({
						code: 'EDIT_UNSUPPORTED_PAGE_SIZE_DEPENDENCY',
					});
					compared++;
					continue;
				}
				const result = await editVsdx(original, [command]);
				const actual = await parseVsdx(result.bytes),
					reference = await parseVsdx(await readFile(item.afterFile));
				const page = actual.pages.find((page) => page.id === command.pageId)!,
					expected = reference.pages.find((page) => page.id === command.pageId)!;
				expect(page.width).toBeCloseTo(expected.width, 12);
				expect(page.height).toBeCloseTo(expected.height, 12);
				expect(page.shapes).toEqual(expected.shapes);
				expect(page.connectors).toEqual(expected.connectors);
				expect(actual.pages.find((page) => page.isBackground)).toEqual(
					reference.pages.find((page) => page.isBackground),
				);
				expect(result.changedParts).toEqual(['visio/pages/pages.xml']);
				const before = await VisioPackage.open(original),
					after = await VisioPackage.open(result.bytes);
				for (const path of before.paths())
					if (!result.changedParts.includes(path))
						expect(await after.readBytes(path)).toEqual(await before.readBytes(path));
				await writeFile(
					join(output, basename(item.beforeFile).replace('-before.vsdx', '-core.vsdx')),
					result.bytes,
				);
				compared++;
			}
			expect(compared).toBe(
				variable.endsWith('FACTORY_DIR') ? 1 : variable.endsWith('MODES_DIR') ? 15 : 12,
			);
		},
		30_000,
	);
}
