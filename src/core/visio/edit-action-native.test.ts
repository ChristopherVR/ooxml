import { readFile } from 'node:fs/promises';
import { join } from 'node:path';
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { editVsdx } from './edit';
import { parseVsdx } from './parser';

// Optional: `scripts/record-visio-actions.ps1 -OutputDirectory <dir>` saves the Annotations
// "Callout - curved" master before and after Visio runs its "Show Box" action. Running the action
// here must save the same cells.
const native = process.env.VISIO_NATIVE_ACTIONS_DIR;
const PAGE = 'visio/pages/page1.xml';

/** The page shapes in one quote style, with the sections in a fixed order (Visio orders them by kind). */
const page = async (bytes: Uint8Array) =>
	(await (await JSZip.loadAsync(bytes)).file(PAGE)!.async('string'))
		.replace(/^[\s\S]*?<Shapes>/, '')
		.replaceAll("'", '"')
		.replace(/(?:<Section [\s\S]*?<\/Section>)+/g, (run) =>
			(run.match(/<Section [\s\S]*?<\/Section>/g) ?? []).sort().join(''),
		);

describe.skipIf(!native)('native Visio actions', () => {
	it('writes what Visio writes when the Show Box action runs', async () => {
		const before = new Uint8Array(await readFile(join(native!, 'action-before.vsdx')));
		const after = new Uint8Array(await readFile(join(native!, 'action-after.vsdx')));
		const shape = (await parseVsdx(before)).pages[0]!.shapes[0]!;
		const box = shape.actions!.find((item) => item.menu === 'Show Box')!;
		expect(box.run.kind).toBe('setf');
		const saved = await editVsdx(before, [
			{ type: 'run-action', pageId: '0', shapeId: shape.id, row: box.row },
		]);
		expect(await page(saved.bytes)).toBe(await page(after));
		expect((await parseVsdx(saved.bytes)).pages[0]!.shapes[0]!.actions![0]!.menu).toBe('Hide Box');
	});
});
