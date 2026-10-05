export type RectanglePoint = readonly [number, number];

/**
 * Rounding is the arc radius at contiguous path segments, in internal inches.
 * This deliberately accepts only one explicitly closed, axis-aligned rectangle.
 * https://learn.microsoft.com/en-us/office/client-developer/visio/rounding-cell-line-format-section
 */
export function roundedRectanglePath(
	points: readonly RectanglePoint[] | undefined,
	rounding: number,
	point: (x: number, y: number) => string,
): string | undefined {
	if (!points || points.length !== 5 || !Number.isFinite(rounding) || rounding <= 0)
		return undefined;
	if (!points.every((p) => p.every(Number.isFinite))) return undefined;
	const first = points[0]!,
		last = points[4]!;
	if (first[0] !== last[0] || first[1] !== last[1]) return undefined;
	const directions: RectanglePoint[] = [],
		lengths: number[] = [];
	for (let i = 0; i < 4; i++) {
		const from = points[i]!,
			to = points[i + 1]!;
		const dx = to[0] - from[0],
			dy = to[1] - from[1];
		if ((dx === 0) === (dy === 0)) return undefined;
		const length = Math.abs(dx || dy);
		directions.push([dx / length, dy / length]);
		lengths.push(length);
	}
	for (let i = 0; i < 4; i++) {
		const a = directions[i]!,
			b = directions[(i + 1) % 4]!;
		if (a[0] * b[0] + a[1] * b[1] !== 0) return undefined;
	}
	const radius = Math.min(rounding, ...lengths.map((length) => length / 2));
	// Path serialization has nanoinch precision; do not assert a zero-radius arc.
	if (radius < 1e-9) return undefined;
	const before: RectanglePoint[] = [],
		after: RectanglePoint[] = [];
	for (let i = 0; i < 4; i++) {
		const vertex = points[i]!,
			incoming = directions[(i + 3) % 4]!,
			outgoing = directions[i]!;
		const a: RectanglePoint = [vertex[0] - incoming[0] * radius, vertex[1] - incoming[1] * radius];
		const b: RectanglePoint = [vertex[0] + outgoing[0] * radius, vertex[1] + outgoing[1] * radius];
		if ((a[0] === vertex[0] && a[1] === vertex[1]) || (b[0] === vertex[0] && b[1] === vertex[1]))
			return undefined;
		before.push(a);
		after.push(b);
	}
	const a = directions[0]!,
		b = directions[1]!;
	const sweep = a[0] * b[1] - a[1] * b[0] > 0 ? 1 : 0;
	const start = after[0]!,
		commands = [`M ${point(start[0], start[1])}`];
	for (let i = 1; i <= 4; i++) {
		const entry = before[i % 4]!,
			exit = after[i % 4]!;
		commands.push(
			`L ${point(entry[0], entry[1])}`,
			`A ${point(radius, radius)} 0 0 ${sweep} ${point(exit[0], exit[1])}`,
		);
	}
	commands.push('Z');
	return commands.join(' ');
}

/**
 * Open orthogonal line chains with enough space for the saved radius, without
 * clamping. Repeated vertices deliberately remain unsupported: Visio authors
 * use them to suppress rounding (see docs/visio-connector-rounding.md).
 */
export function roundedOrthogonalPath(
	points: readonly RectanglePoint[] | undefined,
	radius: number,
	point: (x: number, y: number) => string,
): { path: string; extraCommands: number } | undefined {
	if (!points || points.length < 3 || !Number.isFinite(radius) || radius < 1e-9) return undefined;
	if (!points.every((p) => p.every(Number.isFinite))) return undefined;
	const first = points[0]!,
		last = points[points.length - 1]!;
	if (first[0] === last[0] && first[1] === last[1]) return undefined;
	const directions: RectanglePoint[] = [],
		lengths: number[] = [];
	for (let i = 1; i < points.length; i++) {
		const from = points[i - 1]!,
			to = points[i]!;
		const dx = to[0] - from[0],
			dy = to[1] - from[1];
		if ((dx === 0) === (dy === 0)) return undefined;
		const length = Math.abs(dx || dy);
		if (!Number.isFinite(length)) return undefined;
		directions.push([dx / length, dy / length]);
		lengths.push(length);
	}
	const corners = [false];
	for (let i = 1; i < points.length - 1; i++) {
		const a = directions[i - 1]!,
			b = directions[i]!;
		const dot = a[0] * b[0] + a[1] * b[1];
		if (dot < 0) return undefined; // Reversals are not circular quarter turns.
		corners.push(dot === 0);
	}
	corners.push(false);
	// Each adjacent corner consumes radius along this source segment. Keeping
	// collinear vertices avoids silently changing their available rounding space.
	for (let i = 0; i < lengths.length; i++) {
		const required = (Number(corners[i]) + Number(corners[i + 1])) * radius;
		if (lengths[i]! < required) return undefined;
	}
	const commands = [`M ${point(first[0], first[1])}`];
	let extraCommands = 0;
	for (let i = 1; i < points.length - 1; i++) {
		const vertex = points[i]!;
		if (!corners[i]) {
			commands.push(`L ${point(vertex[0], vertex[1])}`);
			continue;
		}
		const incoming = directions[i - 1]!,
			outgoing = directions[i]!;
		const entry: RectanglePoint = [
			vertex[0] - incoming[0] * radius,
			vertex[1] - incoming[1] * radius,
		];
		const exit: RectanglePoint = [
			vertex[0] + outgoing[0] * radius,
			vertex[1] + outgoing[1] * radius,
		];
		// Refuse cancellation/serialization collapse rather than emit false arcs.
		if (
			!entry.every(Number.isFinite) ||
			!exit.every(Number.isFinite) ||
			point(...entry) === point(...vertex) ||
			point(...exit) === point(...vertex)
		)
			return undefined;
		const sweep = incoming[0] * outgoing[1] - incoming[1] * outgoing[0] > 0 ? 1 : 0;
		commands.push(
			`L ${point(...entry)}`,
			`A ${point(radius, radius)} 0 0 ${sweep} ${point(...exit)}`,
		);
		extraCommands++;
	}
	if (!extraCommands) return undefined;
	commands.push(`L ${point(last[0], last[1])}`);
	return { path: commands.join(' '), extraCommands };
}
