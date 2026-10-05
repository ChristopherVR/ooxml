import { workerData, parentPort } from 'node:worker_threads';
import { createRequire } from 'node:module';
import { pathToFileURL } from 'node:url';
import { LIMITS, preflight } from './preflight.mjs';
import { sanitizeVectorTree } from './sanitize.mjs';

// This is an audit harness for reviewed local test inputs, not an untrusted-code sandbox.
const requireCore = createRequire(new URL('../../../ooxml/package.json', import.meta.url));
const warnings = [];
let warningCount = 0;
console.warn = (...args) => {
	warningCount++;
	if (warnings.length < 64) warnings.push(args.map(String).join(' ').slice(0, 512));
};

try {
	if (workerData.mode === 'watchdog-test') {
		for (;;) {}
	}
	const input = new Uint8Array(workerData.input);
	const scan = preflight(input);
	if (workerData.mode === 'safe' && !scan.eligible) {
		parentPort.postMessage({ status: 'unsupported', scan });
	} else {
		const converter = await import(pathToFileURL(requireCore.resolve('emf-converter')).href);
		const options = {
			maxCanvasDimension: LIMITS.outputDimension,
			maxWidth: 512,
			maxHeight: 512,
			maxRecords: LIMITS.records,
			exactRasterOps: false,
			gdiAntialias: true,
		};
		const tree = await converter.convertMetafileToSvgTree(workerData.input, options);
		if (workerData.mode === 'safe') {
			if (!tree || warningCount) throw new Error('emf.incomplete-output');
			const vector = sanitizeVectorTree(tree);
			parentPort.postMessage({ status: 'ok', vector, scan });
		} else {
			const canvas = requireCore('@napi-rs/canvas');
			globalThis.ImageData = canvas.ImageData;
			const url = await converter.convertMetafileToDataUrl(workerData.input, {
				...options,
				gdiAntialias: false,
			});
			const metrics = url ? await measurePixels(url, canvas) : null;
			const textNodes = [];
			let nodes = 0;
			let chars = 0;
			const visit = (node, depth = 0) => {
				if (!node) return;
				if (++nodes > LIMITS.outputNodes || depth > LIMITS.depth)
					throw new Error('emf.output-limit');
				chars +=
					Object.values(node.attrs).reduce((n, v) => n + String(v).length, 0) +
					(node.text?.length ?? 0);
				if (chars > LIMITS.outputChars) throw new Error('emf.output-limit');
				if (node.tag === 'text')
					textNodes.push({
						text: node.text,
						x: node.attrs.x,
						y: node.attrs.y,
						transform: node.attrs.transform,
					});
				for (const child of node.children ?? []) visit(child, depth + 1);
			};
			visit(tree);
			parentPort.postMessage({
				status: tree ? 'converted' : 'null',
				scan,
				warnings,
				warningCount,
				nodes,
				chars,
				textNodes,
				metrics,
			});
		}
	}
} catch (error) {
	parentPort.postMessage({
		error: error instanceof Error ? error.message : String(error),
		warnings,
		warningCount,
	});
}

async function measurePixels(url, canvas) {
	if (url.length > LIMITS.outputChars * 4) throw new Error('emf.output-limit');
	const image = await canvas.loadImage(url);
	const { width, height } = image;
	if (width * height > 512 * 512) throw new Error('emf.pixel-limit');
	const context = canvas.createCanvas(width, height).getContext('2d');
	context.drawImage(image, 0, 0);
	const rgba = context.getImageData(0, 0, width, height).data;
	let darkRight = 0,
		redLeft = 0,
		redRight = 0,
		minX = width,
		minY = height,
		maxX = -1,
		maxY = -1;
	for (let y = 0; y < height; y++)
		for (let x = 0; x < width; x++) {
			const p = (y * width + x) * 4;
			const [r, g, b, a] = rgba.subarray(p, p + 4);
			if (a < 10) continue;
			if (x >= 90 && r < 120 && g < 120 && b < 120) darkRight++;
			if (r > 160 && g < 90 && b < 90) {
				if (x < 75) redLeft++;
				else redRight++;
				minX = Math.min(x, minX);
				minY = Math.min(y, minY);
				maxX = Math.max(x, maxX);
				maxY = Math.max(y, maxY);
			}
		}
	return {
		width,
		height,
		darkRight,
		redLeft,
		redRight,
		redBox: maxX < 0 ? null : { x: minX, y: minY, width: maxX - minX + 1, height: maxY - minY + 1 },
	};
}
