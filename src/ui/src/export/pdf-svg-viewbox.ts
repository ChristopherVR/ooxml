/** A 2D affine matrix `[a, b, c, d, e, f]`, as the PDF `cm` operator takes it. */
export type SvgMatrix = readonly [number, number, number, number, number, number];

/** Map the SVG's user units into its `width` x `height` box (viewBox and preserveAspectRatio). */
export function viewBoxMatrix(svg: Element, width: number, height: number): SvgMatrix {
	const box = (svg.getAttribute('viewBox') ?? '')
		.trim()
		.split(/[\s,]+/)
		.map(Number);
	if (box.length !== 4 || box.some((v) => !Number.isFinite(v)) || !(box[2]! > 0 && box[3]! > 0))
		return [1, 0, 0, 1, 0, 0];
	const [minX, minY, vw, vh] = box as [number, number, number, number];
	const ratio = (svg.getAttribute('preserveAspectRatio') ?? 'xMidYMid meet').trim();
	let sx = width / vw,
		sy = height / vh;
	if (ratio.startsWith('none')) return [sx, 0, 0, sy, -minX * sx, -minY * sy];
	const slice = ratio.endsWith('slice');
	const s = slice ? Math.max(sx, sy) : Math.min(sx, sy);
	sx = s;
	sy = s;
	const align = ratio.split(/\s+/)[0] ?? 'xMidYMid';
	const dx = width - vw * s,
		dy = height - vh * s;
	const tx = align.includes('xMid') ? dx / 2 : align.includes('xMax') ? dx : 0;
	const ty = align.includes('YMid') ? dy / 2 : align.includes('YMax') ? dy : 0;
	return [s, 0, 0, s, tx - minX * s, ty - minY * s];
}
