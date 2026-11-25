import type { PptxBackgroundRemoval, PptxBackgroundRemovalMark } from 'ooxml-core/pptx';

export interface PictureMarks {
	foreground: readonly PptxBackgroundRemovalMark[];
	background: readonly PptxBackgroundRemovalMark[];
}

export const MAX_PICTURE_MARKS = 128;

/** Flood from the image edge and remove strokes. Keep strokes are barriers. */
export function segmentPicture(
	image: ImageData,
	marks: PictureMarks,
): {
	mask: ImageData;
	removal: PptxBackgroundRemoval;
} {
	const { width, height, data } = image;
	const pixels = width * height;
	const labels = new Uint8Array(pixels);
	const queue = new Uint32Array(pixels);
	let end = 0;
	const enqueue = (index: number) => {
		if (labels[index] !== 0) return;
		labels[index] = 2;
		queue[end++] = index;
	};
	const stroke = (mark: PptxBackgroundRemovalMark, apply: (index: number) => void) => {
		if (![mark.x1, mark.x2, mark.y1, mark.y2].every((v) => Number.isFinite(v) && v >= 0 && v <= 1))
			return;
		const dx = (mark.x2 - mark.x1) * (width - 1);
		const dy = (mark.y2 - mark.y1) * (height - 1);
		const steps = Math.max(1, Math.ceil(Math.hypot(dx, dy)));
		for (let i = 0; i <= steps; i++) {
			const x = Math.round(mark.x1 * (width - 1) + (dx * i) / steps);
			const y = Math.round(mark.y1 * (height - 1) + (dy * i) / steps);
			for (let row = Math.max(0, y - 2); row <= Math.min(height - 1, y + 2); row++) {
				for (let col = Math.max(0, x - 2); col <= Math.min(width - 1, x + 2); col++) {
					apply(row * width + col);
				}
			}
		}
	};
	for (const mark of marks.foreground)
		stroke(mark, (index) => {
			labels[index] = 1;
		});
	for (const mark of marks.background) stroke(mark, enqueue);

	const cornerSamples = [0, width - 1, pixels - width, pixels - 1];
	const distance = (a: number, b: number) =>
		Math.abs(data[a * 4]! - data[b * 4]!) +
		Math.abs(data[a * 4 + 1]! - data[b * 4 + 1]!) +
		Math.abs(data[a * 4 + 2]! - data[b * 4 + 2]!);
	const edge = (index: number) => {
		if (data[index * 4 + 3] === 0 || cornerSamples.some((sample) => distance(sample, index) < 65)) {
			enqueue(index);
		}
	};
	for (let x = 0; x < width; x++) {
		edge(x);
		edge(pixels - width + x);
	}
	for (let y = 1; y < height - 1; y++) {
		edge(y * width);
		edge(y * width + width - 1);
	}

	for (let cursor = 0; cursor < end; cursor++) {
		const current = queue[cursor]!;
		const visit = (next: number) => {
			if (labels[next] === 0 && (data[next * 4 + 3] === 0 || distance(current, next) < 26)) {
				enqueue(next);
			}
		};
		if (current % width > 0) visit(current - 1);
		if (current % width < width - 1) visit(current + 1);
		if (current >= width) visit(current - width);
		if (current < pixels - width) visit(current + width);
	}

	// An unreliable all-background guess leaves the picture visible until the user adds marks.
	const keepWhole = end === pixels && marks.background.length === 0;
	const mask = new ImageData(width, height);
	let left = width,
		right = 0,
		top = height,
		bottom = 0;
	for (let index = 0; index < pixels; index++) {
		if (!keepWhole && labels[index] === 2) continue;
		mask.data[index * 4 + 3] = 255;
		const x = index % width;
		const y = Math.floor(index / width);
		left = Math.min(left, x);
		right = Math.max(right, x + 1);
		top = Math.min(top, y);
		bottom = Math.max(bottom, y + 1);
	}
	return {
		mask,
		removal: {
			left: right ? left / width : 0,
			right: right ? right / width : 1,
			top: bottom ? top / height : 0,
			bottom: bottom ? bottom / height : 1,
			foregroundMarks: [...marks.foreground],
			backgroundMarks: [...marks.background],
		},
	};
}
