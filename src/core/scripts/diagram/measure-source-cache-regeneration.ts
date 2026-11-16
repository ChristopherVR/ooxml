/**
 * Measure the round-trip source-cache writer against the caches Office laid out: for every
 * SmartArt gallery fixture (PowerPoint 16's own diagrams, `pptx/__tests__/fixtures/smartart-gallery`)
 * and the Word fixture, regenerate the drawing from the unedited data, layout, colour and quick-style
 * parts with `buildGeneratedDiagramCache`, then compare each regenerated shape with the Office shape
 * bound to the same presentation point. Reports how many diagrams the writer admits, how many shapes
 * bind to Office's presentation points, and the frame deviation (the worst of x, y, width and height
 * as a fraction of the diagram's larger side).
 *
 * Usage: `bun run scripts/diagram/measure-source-cache-regeneration.ts [--json <file>] [--only <substr>]`
 */
import { readFileSync, readdirSync, writeFileSync } from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import JSZip from 'jszip';

import { buildGeneratedDiagramCache } from '../../diagram/write/cache';
import { parseDiagramDrawing } from '../../diagram/drawing';

const here = path.dirname(fileURLToPath(import.meta.url));
const gallery = path.join(here, '../../pptx/__tests__/fixtures/smartart-gallery');
const word = path.join(here, '../../docx/__fixtures__/smartart-word.docx');
const args = process.argv.slice(2);
const only = args.includes('--only') ? args[args.indexOf('--only') + 1] : undefined;
const jsonOut = args.includes('--json') ? args[args.indexOf('--json') + 1] : undefined;

interface Row {
	file: string;
	admitted: boolean;
	officeShapes: number;
	regenerated: number;
	bound: number;
	maxDeviation?: number;
	meanDeviation?: number;
}

async function measure(file: string, prefix: 'ppt' | 'word'): Promise<Row> {
	const zip = await JSZip.loadAsync(readFileSync(file));
	const read = (name: string) => zip.file(`${prefix}/diagrams/${name}1.xml`)?.async('string');
	const [dataXml, layoutXml, colorsXml, quickStyleXml, drawingXml] = await Promise.all([
		read('data'),
		read('layout'),
		read('colors'),
		read('quickStyle'),
		read('drawing'),
	]);
	const row: Row = {
		file: path.basename(file),
		admitted: false,
		officeShapes: 0,
		regenerated: 0,
		bound: 0,
	};
	if (!dataXml || !layoutXml || !drawingXml) return row;
	const office = parseDiagramDrawing(drawingXml);
	row.officeShapes = office.shapes.length;
	const host =
		prefix === 'ppt'
			? await zip.file('ppt/slides/slide1.xml')?.async('string')
			: await zip.file('word/document.xml')?.async('string');
	const ext = /<(?:a:ext|wp:extent) cx="(\d+)" cy="(\d+)"/u.exec(
		prefix === 'ppt' ? (host?.slice(host.indexOf('graphicFrame')) ?? '') : (host ?? ''),
	);
	if (!ext) return row;
	const extentEmu = { width: Number(ext[1]), height: Number(ext[2]) };
	const layoutId = /uniqueId="([^"]+)"/u.exec(layoutXml)?.[1];
	const cache = buildGeneratedDiagramCache(dataXml, {
		extentEmu,
		source: {
			layoutXml,
			...(colorsXml ? { colorsXml } : {}),
			...(quickStyleXml ? { quickStyleXml } : {}),
		},
		...(layoutId ? { layoutId } : {}),
	});
	if (!cache) return row;
	row.admitted = true;
	const regenerated = parseDiagramDrawing(cache.drawingXml);
	row.regenerated = regenerated.shapes.length;
	const byId = new Map(office.shapes.map((shape) => [shape.modelId, shape]));
	const side = Math.max(extentEmu.width, extentEmu.height);
	const deviations: number[] = [];
	for (const shape of regenerated.shapes) {
		const match = byId.get(shape.modelId);
		if (!match) continue;
		row.bound++;
		const a = shape.frame;
		const b = match.frame;
		deviations.push(
			Math.max(
				Math.abs(a.x - b.x),
				Math.abs(a.y - b.y),
				Math.abs(a.width - b.width),
				Math.abs(a.height - b.height),
			) / side,
		);
	}
	if (deviations.length) {
		row.maxDeviation = Math.max(...deviations);
		row.meanDeviation = deviations.reduce((sum, value) => sum + value, 0) / deviations.length;
	}
	return row;
}

const files = readdirSync(gallery)
	.filter((name) => name.endsWith('.pptx') && (!only || name.includes(only)))
	.map((name) => [path.join(gallery, name), 'ppt'] as const);
if (!only || 'smartart-word.docx'.includes(only)) files.push([word, 'word']);
const rows: Row[] = [];
for (const [file, prefix] of files) {
	try {
		rows.push(await measure(file, prefix));
	} catch (error) {
		rows.push({
			file: path.basename(file),
			admitted: false,
			officeShapes: 0,
			regenerated: 0,
			bound: 0,
		});
		console.error(path.basename(file), error instanceof Error ? error.message : error);
	}
}
const admitted = rows.filter((row) => row.admitted);
const bound = admitted.filter((row) => row.bound > 0);
const within = (limit: number) =>
	bound.filter((row) => (row.maxDeviation ?? Infinity) <= limit).length;
const summary = {
	fixtures: rows.length,
	admitted: admitted.length,
	withBoundShapes: bound.length,
	boundShapes: admitted.reduce((sum, row) => sum + row.bound, 0),
	regeneratedShapes: admitted.reduce((sum, row) => sum + row.regenerated, 0),
	officeShapesOfAdmitted: admitted.reduce((sum, row) => sum + row.officeShapes, 0),
	within1pct: within(0.01),
	within5pct: within(0.05),
	within10pct: within(0.1),
};
console.log(JSON.stringify(summary, null, 2));
for (const row of admitted)
	console.log(
		`${row.file}\tbound ${row.bound}/${row.regenerated} of ${row.officeShapes}\tmax ${((row.maxDeviation ?? NaN) * 100).toFixed(1)}%\tmean ${((row.meanDeviation ?? NaN) * 100).toFixed(1)}%`,
	);
if (jsonOut) writeFileSync(jsonOut, `${JSON.stringify({ summary, rows }, null, '\t')}\n`);
