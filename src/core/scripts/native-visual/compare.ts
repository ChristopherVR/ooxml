// Run from src/core: bun scripts/native-visual/compare.ts <manifest.json>
import { createCanvas, loadImage } from '@napi-rs/canvas';
import { readFileSync, writeFileSync } from 'node:fs';
import { dirname, resolve } from 'node:path';
import { comparePixels } from './metrics';

const manifest = resolve(process.argv[2]!);
const root = dirname(manifest);
const nativeEvidence = JSON.parse(
	readFileSync(resolve(root, 'native.json'), 'utf8').replace(/^\uFEFF/, ''),
);
const spec = JSON.parse(readFileSync(manifest, 'utf8')) as {
	cases: {
		id: string;
		width: number;
		height: number;
		expect?: { maximumError: number; minimumOverlap: number };
	}[];
};
const results = [];
let failed = false;
for (const entry of spec.cases) {
	if (!/^[a-z0-9][a-z0-9-]*$/.test(entry.id)) throw new Error('Invalid case id');
	const capture = nativeEvidence.cases.find((item: { id: string }) => item.id === entry.id);
	if (capture?.nativeSlidePng) {
		const slide = await loadImage(resolve(root, capture.nativeSlidePng));
		const frame = capture.crop;
		if (
			frame.x < 0 ||
			frame.y < 0 ||
			frame.x + frame.width > slide.width ||
			frame.y + frame.height > slide.height
		)
			throw new Error(`${entry.id}: declared native frame extends beyond the slide`);
		const crop = createCanvas(frame.width, frame.height);
		crop
			.getContext('2d')
			.drawImage(
				slide,
				frame.x,
				frame.y,
				frame.width,
				frame.height,
				0,
				0,
				frame.width,
				frame.height,
			);
		writeFileSync(resolve(root, `${entry.id}.native.png`), crop.toBuffer('image/png'));
	}
	const native = await loadImage(resolve(root, `${entry.id}.native.png`));
	const shared = await loadImage(resolve(root, `${entry.id}.shared.png`));
	const width = entry.width;
	const height = entry.height;
	const dimensionsMatch = [native, shared].every(
		(image) => image.width === width && image.height === height,
	);
	// Keep the full export, including any native ink outside the authored shape frame. Size mismatches
	// are reported and never silently accepted. Resampling here only makes the difference view readable.
	const pixels = [native, shared].map((image) => {
		const canvas = createCanvas(width, height);
		const context = canvas.getContext('2d');
		context.drawImage(image, 0, 0, width, height);
		return context.getImageData(0, 0, width, height).data;
	});
	const { difference, ...metrics } = comparePixels(pixels[0]!, pixels[1]!);
	const diff = createCanvas(width, height);
	const diffContext = diff.getContext('2d');
	const imageData = diffContext.createImageData(width, height);
	imageData.data.set(difference);
	diffContext.putImageData(imageData, 0, 0);
	writeFileSync(resolve(root, `${entry.id}.difference.png`), diff.toBuffer('image/png'));
	const contact = createCanvas(width * 3, height + 30);
	const context = contact.getContext('2d');
	context.fillStyle = 'white';
	context.fillRect(0, 0, contact.width, contact.height);
	context.fillStyle = 'black';
	context.font = '14px Arial';
	for (const [index, label] of [
		'Native Office',
		'Shared renderer',
		'Absolute difference x4',
	].entries())
		context.fillText(label, index * width + 8, 20);
	context.drawImage(native, 0, 30, width, height);
	context.drawImage(shared, width, 30, width, height);
	context.drawImage(diff, width * 2, 30);
	writeFileSync(resolve(root, `${entry.id}.comparison.png`), contact.toBuffer('image/png'));
	const contentPresent = metrics.nativeForegroundRatio > 0 && metrics.sharedForegroundRatio > 0;
	const status =
		!dimensionsMatch || !contentPresent
			? 'invalid'
			: entry.expect
				? metrics.normalizedMeanAbsoluteError <= entry.expect.maximumError &&
					metrics.foregroundIntersectionOverUnion >= entry.expect.minimumOverlap
					? 'passed'
					: 'failed'
				: 'measured';
	failed ||= status === 'invalid' || status === 'failed';
	results.push({
		id: entry.id,
		status,
		dimensionsMatch,
		nativeSize: [native.width, native.height],
		sharedSize: [shared.width, shared.height],
		...metrics,
		...(entry.expect ? { expected: entry.expect } : {}),
	});
}
writeFileSync(
	resolve(root, 'comparison.json'),
	JSON.stringify(
		{
			comparedAt: new Date().toISOString(),
			policy:
				'White composite, unregistered full frame, RGB MAE, delta >16 pixel ratio, foreground <245 IoU. Measured is not passed.',
			native: nativeEvidence,
			shared: JSON.parse(readFileSync(resolve(root, 'shared.json'), 'utf8')),
			cases: results,
		},
		null,
		2,
	),
);
console.log(JSON.stringify(results, null, 2));
if (failed) process.exitCode = 1;
