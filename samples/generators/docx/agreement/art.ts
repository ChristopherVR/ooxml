/** The agreement's two small pictures: the co-operative's mark and the logo placeholder. */
import { Canvas, hex, mix } from '../../lib/art';
import { INK } from './styles';

/** The Quillhaven mark: a gold quill across a navy disc. */
export function mark(): Uint8Array {
	const size = 180;
	const canvas = new Canvas(size, size, [255, 255, 255]);
	const navy = hex(INK.navy);
	const gold = hex(INK.gold);
	canvas.circle(90, 90, 86, navy);
	// The vane of the quill, then its shaft and the cut of the nib.
	canvas.polygon(
		[
			[122, 34],
			[140, 52],
			[96, 112],
			[66, 124],
			[72, 94],
		],
		gold,
	);
	canvas.line(134, 42, 52, 138, 5, mix(gold, [255, 255, 255], 0.55));
	canvas.polygon(
		[
			[56, 126],
			[66, 134],
			[46, 146],
		],
		[255, 255, 255],
	);
	return canvas.toPng();
}

/** What the picture control shows until the client drops a logo on it: a frame with a hill and sun. */
export function logoPlaceholder(): Uint8Array {
	const width = 360;
	const height = 220;
	const paper = hex(INK.paper);
	const line = hex(INK.line);
	const canvas = new Canvas(width, height, paper);
	// A dashed frame.
	for (let x = 8; x < width - 8; x += 22) {
		canvas.rect(x, 8, 12, 3, line);
		canvas.rect(x, height - 11, 12, 3, line);
	}
	for (let y = 8; y < height - 8; y += 22) {
		canvas.rect(8, y, 3, 12, line);
		canvas.rect(width - 11, y, 3, 12, line);
	}
	const soft = mix(line, hex(INK.slate), 0.25);
	canvas.circle(250, 76, 20, soft);
	canvas.polygon(
		[
			[70, 172],
			[150, 82],
			[206, 142],
			[236, 116],
			[292, 172],
		],
		soft,
	);
	return canvas.toPng();
}
