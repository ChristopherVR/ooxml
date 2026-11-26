import { encodeBase64 } from '../digest/base64';
import { child } from './sheet';
import { visioOutlineShape } from './stencil-shapes';
import type { VisioStencilMaster } from './stencils';

// Single-image 32px ICO with a 4-bit DIB and a 1-bit AND mask, as Visio saves in Master/Icon.
const SIZE = 32;
const PIXELS = 22 + 40 + 16 * 4;
const MASK = PIXELS + (SIZE * SIZE) / 2;

export function ensureStencilIcon(entry: Element, master?: VisioStencilMaster): void {
	if (child(entry, 'Icon')) return;
	const bytes = new Uint8Array(MASK + (SIZE * SIZE) / 8);
	const view = new DataView(bytes.buffer);
	view.setUint16(2, 1, true);
	view.setUint16(4, 1, true);
	bytes.set([SIZE, SIZE, 16], 6);
	view.setUint16(10, 1, true);
	view.setUint16(12, 4, true);
	view.setUint32(14, bytes.length - 22, true);
	view.setUint32(18, 22, true);
	view.setUint32(22, 40, true);
	view.setInt32(26, SIZE, true);
	view.setInt32(30, SIZE * 2, true);
	view.setUint16(34, 1, true);
	view.setUint16(36, 4, true);
	view.setUint32(42, bytes.length - PIXELS, true);
	bytes.set([160, 97, 35, 0, 243, 215, 177, 0], 66);
	const inside = iconOutline(master);
	const grid = Array.from({ length: SIZE }, (_, y) =>
		Array.from({ length: SIZE }, (_, x) => inside(x + 0.5, y + 0.5)),
	);
	bytes.fill(0xff, MASK);
	for (const [y, row] of grid.entries()) {
		for (const [x, on] of row.entries()) {
			if (!on) continue;
			const fill = row[x - 1] && row[x + 1] && grid[y - 1]?.[x] && grid[y + 1]?.[x];
			const index = PIXELS + y * (SIZE / 2) + (x >> 1);
			bytes[index] = (bytes[index] ?? 0) | ((fill ? 2 : 1) << (x % 2 ? 0 : 4));
			const alpha = MASK + y * (SIZE / 8) + (x >> 3);
			bytes[alpha] = (bytes[alpha] ?? 0) & ~(0x80 >> (x % 8));
		}
	}
	const icon = entry.ownerDocument!.createElementNS(entry.namespaceURI, 'Icon');
	icon.textContent = encodeBase64(bytes);
	entry.insertBefore(icon, child(entry, 'Rel') ?? null);
}

function iconOutline(master?: VisioStencilMaster): (x: number, y: number) => boolean {
	const size = master?.size ?? { width: 1, height: 0.75 };
	const scale = 24 / Math.max(size.width, size.height);
	const w = size.width * scale,
		h = size.height * scale;
	const paths =
		master?.shape === 'ellipse' ? [] : visioOutlineShape(master?.shape ?? 'rectangle').paths;
	return (px, py) => {
		const x = (px - (SIZE - w) / 2) / w,
			y = (py - (SIZE - h) / 2) / h;
		if (master?.shape === 'ellipse') return (x - 0.5) ** 2 + (y - 0.5) ** 2 < 0.25;
		return paths.some((points) => {
			let inside = false;
			let previous = points.at(-1);
			for (const next of points) {
				if (
					previous &&
					next[1] > y !== previous[1] > y &&
					x < ((previous[0] - next[0]) * (y - next[1])) / (previous[1] - next[1]) + next[0]
				)
					inside = !inside;
				previous = next;
			}
			return inside;
		});
	};
}
