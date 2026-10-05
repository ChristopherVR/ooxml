import { createHash } from 'node:crypto';
import { createRequire } from 'node:module';
import { mkdir, readFile, readdir, writeFile } from 'node:fs/promises';
import path from 'node:path';
import JSZip from 'jszip';
import { syntheticCases } from './fixtures.mjs';
import { runIsolated } from './isolated.mjs';
import { LIMITS, preflight } from './preflight.mjs';

const [corpus, output] = process.argv.slice(2);
if (!corpus || !output || process.argv.length !== 4) {
	throw new Error('Usage: node scripts/emf-audit/run.mjs <corpus-directory> <output-directory>');
}
const core = createRequire(new URL('../../../ooxml/package.json', import.meta.url));
const converterEntry = core.resolve('emf-converter');
const packageDir = path.dirname(path.dirname(converterEntry));
const manifest = JSON.parse(await readFile(path.join(packageDir, 'package.json'), 'utf8'));
if (manifest.version !== '3.5.1')
	throw new Error(
		'This characterization harness requires emf-converter 3.5.1. Re-audit before changing the pin.',
	);
const sha = (bytes) => createHash('sha256').update(bytes).digest('hex');
const evidence = {
	generatedAt: new Date().toISOString(),
	environment: {
		node: process.version,
		platform: process.platform,
		canvas: core('@napi-rs/canvas/package.json').version,
	},
	converter: {
		version: manifest.version,
		license: manifest.license,
		cjsSha256: sha(await readFile(converterEntry)),
		esmSha256: sha(await readFile(path.join(packageDir, 'dist/index.mjs'))),
	},
	limits: LIMITS,
	note: 'Audit conversions bypass subset eligibility only for the named reviewed local fixtures. No Windows or browser parity is asserted.',
	synthetic: {},
	corpus: [],
};

for (const [name, bytes] of Object.entries(syntheticCases())) {
	evidence.synthetic[name] = await convert(name, bytes);
}
const badSignature = Buffer.from(syntheticCases()['plain-rectangle']);
badSignature.writeUInt32LE(0, 40);
evidence.synthetic['invalid-signature'] = await convert('invalid-signature', badSignature);
const noEof = syntheticCases()['plain-rectangle'].subarray(0, -20);
evidence.synthetic['missing-eof'] = await convert('missing-eof', noEof);
const unknown = Buffer.from(syntheticCases()['plain-rectangle']);
unknown.writeUInt32LE(0x12345678, 108);
evidence.synthetic['unknown-record'] = await convert('unknown-record', unknown);

for (const source of ['apache-poi', 'libvisio']) {
	const fixtures = path.join(corpus, source, 'fixtures');
	for (const filename of (await readdir(fixtures))
		.filter((name) => name.endsWith('.vsdx'))
		.sort()) {
		const packed = await readFile(path.join(fixtures, filename));
		if (packed.length > 32 * 1024 * 1024)
			throw new Error(`Archive exceeds audit input limit: ${filename}`);
		const zip = await JSZip.loadAsync(packed);
		for (const member of Object.values(zip.files)) {
			if (!/^visio\/media\/[^/]+\.emf$/i.test(member.name)) continue;
			// The audited ZIP library records uncompressed size before materialization.
			if (!member._data || member._data.uncompressedSize > LIMITS.inputBytes)
				throw new Error(`Media exceeds limit: ${member.name}`);
			const bytes = await member.async('uint8array');
			evidence.corpus.push({
				source,
				document: filename,
				part: member.name,
				kind: 'embedded-media',
				...(await convert(`${source}/${filename}:${member.name}`, bytes)),
			});
		}
	}
	const previews = path.join(corpus, source, 'embedded-previews');
	for (const filename of (await readdir(previews)).filter((name) => name.endsWith('.emf')).sort()) {
		const bytes = await readFile(path.join(previews, filename));
		const scan = preflight(bytes);
		evidence.corpus.push({
			source,
			document: filename,
			kind: 'thumbnail-preview',
			bytes: bytes.length,
			sha256: sha(bytes),
			scan,
		});
	}
}
await mkdir(output, { recursive: true });
await writeFile(path.join(output, 'evidence.json'), JSON.stringify(evidence, null, 2) + '\n');
console.log(`Evidence saved to ${path.join(output, 'evidence.json')}`);
console.log(
	JSON.stringify(
		{
			embeddedMedia: evidence.corpus
				.filter((x) => x.kind === 'embedded-media')
				.map((x) => ({
					part: x.part,
					status: x.status,
					eligible: x.scan.eligible,
					warnings: x.warnings,
					nodes: x.nodes,
				})),
			thumbnailPreviews: evidence.corpus.filter((x) => x.kind === 'thumbnail-preview').length,
		},
		null,
		2,
	),
);

async function convert(name, bytes) {
	const start = performance.now();
	try {
		const result = await runIsolated(bytes);
		console.log(`${name}: ${result.status}, ${Math.round(performance.now() - start)}ms`);
		return {
			bytes: bytes.byteLength,
			sha256: sha(bytes),
			durationMs: Math.round(performance.now() - start),
			...result,
		};
	} catch (error) {
		return {
			bytes: bytes.byteLength,
			sha256: sha(bytes),
			scan: preflight(bytes),
			status: 'failed',
			error: error.message,
		};
	}
}
