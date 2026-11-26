// Run after `bun run build`: node src/core/scripts/bench-pptx.mjs [fixture ...]
// For a CPU profile, add --cpu-prof before the script path. PPTX_BENCH_PHASE can
// select load or save, PPTX_BENCH_RUNS sets the sample count, and PPTX_BENCH_MODULE
// selects an alternate built entry point for comparisons.
import { createHash } from 'node:crypto';
import { readFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import { pathToFileURL } from 'node:url';

const defaultFixtures = [
	'e2e/Image_JPG_PNG_Audio_M4_A_Video_MP_4_12_Slides_36_8_MB_ff1095731b.pptx',
	'e2e/36_Slides_Extra_Large_22_5_MB_578ce6bbf3.pptx',
	'e2e/issue-132-hr-deck.pptx',
	'e2e/three-d-parity/three-d-charts.pptx',
	'e2e/chart-gallery.pptx',
	'corpus/smartart-chart-table-mix.pptx',
];
const fixtures = process.argv.length > 2 ? process.argv.slice(2) : defaultFixtures;
const runs = Number(process.env.PPTX_BENCH_RUNS ?? 5);
const phase = process.env.PPTX_BENCH_PHASE ?? 'both';
if (!Number.isSafeInteger(runs) || runs < 1) throw new Error('Expected a positive run count');
if (!['both', 'load', 'save'].includes(phase)) throw new Error('Expected both, load or save');

const entry = process.env.PPTX_BENCH_MODULE;
const { PptxHandler } = await import(
	entry
		? pathToFileURL(resolve(entry)).href
		: new URL('../dist/pptx/index.mjs', import.meta.url).href
);
const fixedDate = new Date('2000-01-01T12:00:00Z');

async function load(handler, bytes) {
	return await handler.load(bytes);
}

async function save(handler, slides) {
	return await handler.save(slides, { fixedDate });
}

async function measure(samples, operation) {
	const start = performance.now();
	const cpuStart = process.cpuUsage();
	const result = await operation();
	const cpu = process.cpuUsage(cpuStart);
	samples.push({ wall: performance.now() - start, cpu: (cpu.user + cpu.system) / 1000 });
	return result;
}

function statistics(samples, key) {
	const values = samples.map((sample) => sample[key]).sort((a, b) => a - b);
	if (values.length === 0) return undefined;
	const low = values[Math.floor((values.length - 1) / 2)];
	const high = values[Math.floor(values.length / 2)];
	const round = (value) => Number(value.toFixed(1));
	return { median: round((low + high) / 2), min: round(values[0]), max: round(values.at(-1)) };
}

for (const fixture of fixtures) {
	const bytes = new Uint8Array(
		await readFile(new URL(`../pptx/__tests__/fixtures/${fixture}`, import.meta.url)),
	).buffer;
	const loadSamples = [];
	const saveSamples = [];
	let savedHash;
	let savedBytes;
	let slideCount;
	for (let i = 0; i <= runs; i++) {
		const handler = new PptxHandler();
		try {
			const data =
				i > 0 && phase !== 'save'
					? await measure(loadSamples, () => load(handler, bytes))
					: await load(handler, bytes);
			slideCount = data.slides.length;
			if (!slideCount) throw new Error(`No slides loaded from ${fixture}`);
			if (phase === 'load') continue;
			const result =
				i > 0
					? await measure(saveSamples, () => save(handler, data.slides))
					: await save(handler, data.slides);
			const hash = createHash('sha256').update(result).digest('hex');
			if (savedHash && savedHash !== hash) throw new Error(`Save bytes varied for ${fixture}`);
			savedHash = hash;
			savedBytes = result.byteLength;
		} finally {
			handler.dispose();
		}
	}
	console.log(
		JSON.stringify({
			fixture,
			runs,
			slideCount,
			savedBytes,
			savedHash,
			loadMs: statistics(loadSamples, 'wall'),
			loadCpuMs: statistics(loadSamples, 'cpu'),
			saveMs: statistics(saveSamples, 'wall'),
			saveCpuMs: statistics(saveSamples, 'cpu'),
		}),
	);
}
