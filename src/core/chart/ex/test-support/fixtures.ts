// The committed chartex parts, shared by the chartex round-trip, render and parity tests.
import { readFileSync } from 'node:fs';
import path from 'node:path';
import JSZip from 'jszip';

const root = path.join(import.meta.dirname, '..');

/** Every committed package holding a chartex part (relative to `chart/ex`), with its part count. */
export const CHART_EX_FIXTURES: [string, number][] = [
	['../../pptx/__tests__/fixtures/e2e/chart-gallery.pptx', 7],
	['../../docx/__fixtures__/chart-families.docx', 4],
	['../../xlsx/__fixtures__/excel-chart-families.xlsx', 4],
];

/** The chartex parts of a package: `[part name, xml]`. */
export async function chartExParts(file: string): Promise<[string, string][]> {
	const zip = await JSZip.loadAsync(readFileSync(path.join(root, file)));
	const names = Object.keys(zip.files).filter((name) => /charts\/chart(Ex)?\d+\.xml$/.test(name));
	const parts = await Promise.all(
		names.map(async (name) => [name, await zip.file(name)!.async('string')] as [string, string]),
	);
	return parts.filter(([, xml]) => /<cx:chartSpace/.test(xml));
}
