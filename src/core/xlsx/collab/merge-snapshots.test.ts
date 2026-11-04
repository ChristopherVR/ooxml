import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { getCell, putCell, deleteCell } from '../cells';
import { loadXlsx } from '../read/load';
import { saveXlsx } from '../write/save';
import { createWorkbook } from '../workbook';
import { mergeWorkbookSnapshots } from './merge-snapshots';

async function branches() {
	const workbook = createWorkbook();
	putCell(workbook.sheets[0]!, 0, 0, { value: 'base A' });
	putCell(workbook.sheets[0]!, 0, 1, { value: 'base B' });
	const base = await saveXlsx(workbook);
	return { base, local: await loadXlsx(base), remote: await loadXlsx(base) };
}

describe('saved workbook merge', () => {
	it('merges independent string edits after each writer regenerates shared-string indexes', async () => {
		const { base, local, remote } = await branches();
		putCell(local.sheets[0]!, 0, 0, { value: 'local A' });
		putCell(remote.sheets[0]!, 0, 1, { value: 'remote B' });
		const result = await mergeWorkbookSnapshots(
			base,
			await saveXlsx(local),
			await saveXlsx(remote),
		);
		expect(result.status).toBe('merged');
		if (result.status !== 'merged') return;
		const book = await loadXlsx(result.bytes);
		expect(getCell(book.sheets[0]!, 0, 0)?.value).toBe('local A');
		expect(getCell(book.sheets[0]!, 0, 1)?.value).toBe('remote B');
	});
	it('combines a value and an independent style change to the same existing cell', async () => {
		const { base, local, remote } = await branches();
		getCell(local.sheets[0]!, 0, 0)!.value = 'new value';
		const style = structuredClone(remote.styles[0]!);
		style.font.bold = true;
		remote.styles.push(style);
		getCell(remote.sheets[0]!, 0, 0)!.styleId = remote.styles.length - 1;
		const result = await mergeWorkbookSnapshots(
			base,
			await saveXlsx(local),
			await saveXlsx(remote),
		);
		expect(result.status).toBe('merged');
		if (result.status !== 'merged') return;
		const book = await loadXlsx(result.bytes),
			cell = getCell(book.sheets[0]!, 0, 0)!;
		expect(cell.value).toBe('new value');
		expect(book.styles[cell.styleId ?? 0]?.font.bold).toBe(true);
	});
	it('preserves independent additions and deletions while refusing same-cell conflicts', async () => {
		const { base, local, remote } = await branches();
		deleteCell(local.sheets[0]!, 0, 0);
		putCell(remote.sheets[0]!, 1, 0, { value: 'added' });
		const result = await mergeWorkbookSnapshots(
			base,
			await saveXlsx(local),
			await saveXlsx(remote),
		);
		expect(result.status).toBe('merged');
		if (result.status !== 'merged') return;
		const book = await loadXlsx(result.bytes);
		expect(getCell(book.sheets[0]!, 0, 0)).toBeUndefined();
		expect(getCell(book.sheets[0]!, 1, 0)?.value).toBe('added');
		getCell(remote.sheets[0]!, 0, 0)!.value = 'edited deleted cell';
		expect(
			(await mergeWorkbookSnapshots(base, await saveXlsx(local), await saveXlsx(remote))).status,
		).toBe('unresolved');
	});
	it('retains unilaterally changed unshared model fields and opaque referenced binary parts', async () => {
		const { base, local, remote } = await branches();
		local.sheets[0]!.pageSetup = { orientation: 'landscape' };
		putCell(remote.sheets[0]!, 0, 1, { value: 'remote' });
		const link = '<Relationship Id="opaque" Type="urn:test" Target="custom/preserved.bin"/>';
		const patch = async (bytes: Uint8Array, value: number) => {
			const zip = await JSZip.loadAsync(bytes);
			zip.file(
				'_rels/.rels',
				(await zip.file('_rels/.rels')!.async('string')).replace(
					'</Relationships>',
					link + '</Relationships>',
				),
			);
			zip.file('custom/preserved.bin', new Uint8Array([value]));
			return zip.generateAsync({ type: 'uint8array' });
		};
		const result = await mergeWorkbookSnapshots(
			await patch(base, 1),
			await patch(await saveXlsx(local), 2),
			await patch(await saveXlsx(remote), 1),
		);
		expect(result.status).toBe('merged');
		if (result.status !== 'merged') return;
		const book = await loadXlsx(result.bytes);
		expect(book.sheets[0]?.pageSetup?.orientation).toBe('landscape');
		expect(
			Array.from(
				await (
					await JSZip.loadAsync(result.bytes)
				)
					.file('custom/preserved.bin')!
					.async('uint8array'),
			),
		).toEqual([2]);
	});
	it.each(['xl/workbook.xml', 'xl/worksheets/sheet1.xml'])(
		'does not silently take one branch for an unmapped source extension change in %s',
		async (path) => {
			const { base, local, remote } = await branches();
			putCell(remote.sheets[0]!, 0, 1, { value: 'remote' });
			const zip = await JSZip.loadAsync(await saveXlsx(local));
			zip.file(
				path,
				(await zip.file(path)!.async('string')).replace(
					/(<\/(?:workbook|worksheet)>)/u,
					'<extLst><ext uri="urn:new"/></extLst>$1',
				),
			);
			const result = await mergeWorkbookSnapshots(
				base,
				await zip.generateAsync({ type: 'uint8array' }),
				await saveXlsx(remote),
			);
			expect(result.status).toBe('unresolved');
			if (result.status === 'unresolved')
				expect(result.issues.join()).toContain('Unmapped source XML');
		},
	);
	it('normalizes independently allocated style ids on different cells', async () => {
		const { base, local, remote } = await branches();
		const bold = structuredClone(local.styles[0]!);
		bold.font.bold = true;
		const italic = structuredClone(remote.styles[0]!);
		italic.font.italic = true;
		local.styles.push(bold);
		remote.styles.push(italic);
		getCell(local.sheets[0]!, 0, 0)!.styleId = 1;
		getCell(remote.sheets[0]!, 0, 1)!.styleId = 1;
		const result = await mergeWorkbookSnapshots(
			base,
			await saveXlsx(local),
			await saveXlsx(remote),
		);
		expect(result.status).toBe('merged');
		if (result.status !== 'merged') return;
		const book = await loadXlsx(result.bytes);
		expect(book.styles[getCell(book.sheets[0]!, 0, 0)?.styleId ?? 0]?.font.bold).toBe(true);
		expect(book.styles[getCell(book.sheets[0]!, 0, 1)?.styleId ?? 0]?.font.italic).toBe(true);
	});
	it('does not treat conflicting formula caches as user edits and preserves unmapped cell metadata', async () => {
		const { local: initial } = await branches();
		putCell(initial.sheets[0]!, 1, 0, { value: 2, formula: '1+1' });
		const base = await saveXlsx(initial),
			local = await loadXlsx(base),
			remote = await loadXlsx(base);
		getCell(local.sheets[0]!, 0, 0)!.value = 'local';
		getCell(remote.sheets[0]!, 0, 1)!.value = 'remote';
		getCell(local.sheets[0]!, 1, 0)!.value = 3;
		getCell(remote.sheets[0]!, 1, 0)!.value = 4;
		const result = await mergeWorkbookSnapshots(
			base,
			await saveXlsx(local),
			await saveXlsx(remote),
		);
		expect(result.status).toBe('merged');
		if (result.status !== 'merged') return;
		const book = await loadXlsx(result.bytes);
		expect(getCell(book.sheets[0]!, 1, 0)?.formula).toBe('1+1');
		expect(book.fullCalcOnLoad).toBe(true);
	});
	it('retains a phonetic guide when another peer formats the same cell', async () => {
		const { local: initial } = await branches();
		getCell(initial.sheets[0]!, 0, 0)!.phonetic = { runs: [{ sb: 0, eb: 1, text: 'reading' }] };
		const base = await saveXlsx(initial),
			local = await loadXlsx(base),
			remote = await loadXlsx(base);
		getCell(local.sheets[0]!, 0, 1)!.value = 'local';
		const bold = structuredClone(remote.styles[0]!);
		bold.font.bold = true;
		remote.styles.push(bold);
		getCell(remote.sheets[0]!, 0, 0)!.styleId = 1;
		const result = await mergeWorkbookSnapshots(
			base,
			await saveXlsx(local),
			await saveXlsx(remote),
		);
		expect(result.status).toBe('merged');
		if (result.status !== 'merged') return;
		expect(getCell((await loadXlsx(result.bytes)).sheets[0]!, 0, 0)?.phonetic?.runs[0]?.text).toBe(
			'reading',
		);
	});
	it('uses remote save stamps while still rejecting conflicting meaningful document properties', async () => {
		const { local: initial } = await branches();
		initial.properties.modified = '2026-01-01T00:00:00Z';
		initial.properties.lastModifiedBy = 'base';
		initial.properties.title = 'baseline title';
		const base = await saveXlsx(initial),
			local = await loadXlsx(base),
			remote = await loadXlsx(base);
		local.properties.modified = '2026-01-02T00:00:00Z';
		local.properties.lastModifiedBy = 'local';
		remote.properties.modified = '2026-01-03T00:00:00Z';
		remote.properties.lastModifiedBy = 'remote';
		getCell(local.sheets[0]!, 0, 0)!.value = 'local';
		getCell(remote.sheets[0]!, 0, 1)!.value = 'remote';
		const result = await mergeWorkbookSnapshots(
			base,
			await saveXlsx(local),
			await saveXlsx(remote),
		);
		expect(result.status).toBe('merged');
		if (result.status !== 'merged') return;
		const book = await loadXlsx(result.bytes);
		expect(book.properties.modified).toBe(remote.properties.modified);
		expect(book.properties.lastModifiedBy).toBe('remote');
		local.properties.title = 'local title';
		remote.properties.title = 'remote title';
		expect(
			(await mergeWorkbookSnapshots(base, await saveXlsx(local), await saveXlsx(remote))).status,
		).toBe('unresolved');
	});
});
