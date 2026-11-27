/**
 * Builds the sample documents and `samples/manifest.json` from the generators.
 *
 *   bun samples/generators/build.ts                 # write every sample and the manifest
 *   bun samples/generators/build.ts --only pptx     # write one product's files (no manifest)
 *   bun samples/generators/build.ts --sample <id>   # write one file (no manifest)
 *   bun samples/generators/build.ts --check         # regenerate in memory and compare
 *   bun samples/generators/build.ts --determinism   # generate everything twice and compare
 *
 * `--check` fails when a committed file, or the manifest, differs from what the generators
 * produce now: run it after changing a generator or a core writer. `--determinism` writes nothing
 * and fails when two runs of a generator disagree; `scripts/samples.test.mjs` runs it. Thumbnails
 * are rendered separately (`samples/generators/thumbnails.ps1`), so the manifest only names where
 * each one lives.
 */
import { createHash } from 'node:crypto';
import { existsSync } from 'node:fs';
import { mkdir, readFile, writeFile } from 'node:fs/promises';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { normalizePackage } from './lib/package';
import type { SampleDefinition, SampleProduct } from './lib/types';

const ROOT = resolve(dirname(fileURLToPath(import.meta.url)), '..');
/** U+2014, which the repository's style bans; built from its code so this file holds none. */
const EM_DASH = String.fromCharCode(0x2014);
const PRODUCTS: readonly SampleProduct[] = ['docx', 'xlsx', 'pptx', 'visio'];
const EXTENSIONS: Record<SampleProduct, string> = {
	docx: '.docx',
	xlsx: '.xlsx',
	pptx: '.pptx',
	visio: '.vsdx',
};

export interface ManifestEntry {
	id: string;
	product: SampleProduct;
	title: string;
	description: string;
	features: string[];
	/** Path of the file, relative to `samples/`. */
	file: string;
	/** Size of the file in bytes. */
	size: number;
	sha256: string;
	/** Path of the thumbnail, relative to `samples/`. */
	thumbnail: string;
}

const argv = process.argv.slice(2);
const value = (flag: string) => (argv.includes(flag) ? argv[argv.indexOf(flag) + 1] : undefined);
const check = argv.includes('--check');
const determinism = argv.includes('--determinism');
const only = value('--only')?.split(',');
const single = value('--sample');
const partial = Boolean(only || single) || determinism;

async function definitions(): Promise<SampleDefinition[]> {
	const all: SampleDefinition[] = [];
	for (const product of PRODUCTS) {
		if (only && !only.includes(product)) continue;
		const module: { samples: SampleDefinition[] } = await import(`./${product}/index`);
		all.push(...module.samples);
	}
	return single ? all.filter((sample) => sample.id === single) : all;
}

function validate(sample: SampleDefinition, seen: Set<string>): void {
	const fail = (why: string) => {
		throw new Error(`${sample.id}: ${why}`);
	};
	if (!/^[a-z0-9]+(?:-[a-z0-9]+)*$/u.test(sample.id)) fail('the id must be a lower-case slug');
	if (seen.has(sample.id)) fail('the id is used twice');
	seen.add(sample.id);
	if (sample.fileName !== `${sample.id}${EXTENSIONS[sample.product]}`)
		fail(`the file must be named ${sample.id}${EXTENSIONS[sample.product]}`);
	if (!sample.title || !sample.description) fail('title and description are required');
	if (sample.features.length < 3) fail('list at least three features');
	const text = [sample.title, sample.description, ...sample.features].join(' ');
	if (text.includes(EM_DASH)) fail('no em-dashes in sample copy');
}

const samples = await definitions();
if (samples.length === 0) throw new Error('No sample matches the given filter.');
const seen = new Set<string>();
const manifest: ManifestEntry[] = [];
const problems: string[] = [];

for (const sample of samples) {
	validate(sample, seen);
	const started = performance.now();
	const bytes = await normalizePackage(await sample.generate());
	const file = `${sample.product}/${sample.fileName}`;
	const path = join(ROOT, file);
	const entry: ManifestEntry = {
		id: sample.id,
		product: sample.product,
		title: sample.title,
		description: sample.description,
		features: sample.features,
		file,
		size: bytes.length,
		sha256: createHash('sha256').update(bytes).digest('hex'),
		thumbnail: `thumbnails/${sample.id}.webp`,
	};
	manifest.push(entry);
	if (determinism) {
		const again = await normalizePackage(await sample.generate());
		if (createHash('sha256').update(again).digest('hex') !== entry.sha256)
			problems.push(`${file}: two runs of its generator gave different bytes`);
	} else if (check) {
		const committed = existsSync(path) ? await readFile(path) : undefined;
		if (!committed) problems.push(`${file} is missing`);
		else if (!committed.equals(bytes))
			problems.push(`${file} differs from what its generator writes now`);
		if (!existsSync(join(ROOT, entry.thumbnail))) problems.push(`${entry.thumbnail} is missing`);
	} else {
		await mkdir(dirname(path), { recursive: true });
		await writeFile(path, bytes);
	}
	const seconds = ((performance.now() - started) / 1000).toFixed(1);
	const did = determinism ? 'generated twice' : check ? 'checked' : 'wrote';
	console.log(`${did} ${file} (${bytes.length} bytes, ${seconds} s)`);
}

const manifestPath = join(ROOT, 'manifest.json');
const manifestText = `${JSON.stringify(
	{
		note: 'Generated by samples/generators/build.ts. Do not edit by hand.',
		samples: manifest,
	},
	null,
	'\t',
)}\n`;
if (!partial) {
	if (check) {
		const committed = existsSync(manifestPath) ? await readFile(manifestPath, 'utf8') : '';
		if (committed !== manifestText) problems.push('manifest.json does not match the generators');
	} else {
		await writeFile(manifestPath, manifestText);
	}
}

if (problems.length > 0) {
	console.error(
		`Samples failed the check:\n${problems.map((problem) => `  - ${problem}`).join('\n')}`,
	);
	if (check) console.error('Run `bun samples/generators/build.ts` and commit the result.');
	process.exit(1);
}
