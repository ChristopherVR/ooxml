/** Affine matrices as PDF writes them: `[a b c d e f]`. */
export type Matrix = [number, number, number, number, number, number];

export const multiply = (m: Matrix, n: Matrix): Matrix => [
	m[0] * n[0] + m[1] * n[2],
	m[0] * n[1] + m[1] * n[3],
	m[2] * n[0] + m[3] * n[2],
	m[2] * n[1] + m[3] * n[3],
	m[4] * n[0] + m[5] * n[2] + n[4],
	m[4] * n[1] + m[5] * n[3] + n[5],
];

export const apply = (m: Matrix, x: number, y: number): [number, number] => [
	m[0] * x + m[2] * y + m[4],
	m[1] * x + m[3] * y + m[5],
];

/** The matrix from user space to reading space: crop box origin, rotation, y down. */
export function pageMatrix(node: { cropBox: number[]; rotate: number }): {
	matrix: Matrix;
	width: number;
	height: number;
} {
	const [x0, y0, x1, y1] = node.cropBox as [number, number, number, number];
	const w = x1 - x0;
	const h = y1 - y0;
	switch (node.rotate) {
		case 90:
			return { matrix: [0, 1, 1, 0, -y0, -x0], width: h, height: w };
		case 180:
			return { matrix: [-1, 0, 0, 1, x1, -y0], width: w, height: h };
		case 270:
			return { matrix: [0, -1, -1, 0, y1, x1], width: h, height: w };
		default:
			return { matrix: [1, 0, 0, -1, -x0, y1], width: w, height: h };
	}
}
