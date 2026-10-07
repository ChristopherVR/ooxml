import type { VisioShape } from '../model';

/** Visual handle eligibility only; source-backed core admission remains authoritative. */
export function visioStraightLineHandles(shape: VisioShape): boolean {
	if (
		shape.kind !== 'connector' ||
		shape.height !== 0 ||
		!(shape.width > 0) ||
		!Number.isFinite(shape.width) ||
		shape.children.length ||
		shape.masterId ||
		shape.hidden ||
		shape.geometry.length !== 1
	)
		return false;
	const geometry = shape.geometry[0]!;
	const match = /^M\s+0\s+0\s+L\s+([-+\d.eE]+)\s+0$/.exec(geometry.path.trim());
	return (
		!geometry.fill && geometry.stroke && !!match && Math.abs(Number(match[1]) - shape.width) <= 1e-6
	);
}
