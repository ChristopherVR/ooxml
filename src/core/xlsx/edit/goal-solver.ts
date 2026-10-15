// The numeric core of Goal Seek: finds x with f(x) = goal by a secant search kept inside a
// sign-changing bracket (bisection when a secant step leaves it). Defaults are Excel's: 100
// iterations and a maximum change of 0.001.

export interface GoalSolverOptions {
	/** Most evaluations after the starting one (default 100). */
	maxIterations?: number;
	/** The result counts as the goal when it is within this distance of it (default 0.001). */
	maxChange?: number;
}

export interface GoalSolution {
	/** Whether the result came within `maxChange` of the goal. */
	found: boolean;
	/** The best input value found. */
	x: number;
	/** `f(x)` at that input. */
	value: number;
	iterations: number;
}

export const DEFAULT_GOAL_ITERATIONS = 100;
export const DEFAULT_GOAL_MAX_CHANGE = 0.001;

interface Point {
	x: number;
	/** `f(x) - goal`. */
	y: number;
}

/** Solves `f(x) = goal` from the starting guess `start`; `f` may return NaN for an error. */
export function solveGoal(
	f: (x: number) => number,
	goal: number,
	start: number,
	options: GoalSolverOptions = {},
): GoalSolution {
	const maxIterations = options.maxIterations ?? DEFAULT_GOAL_ITERATIONS;
	const tolerance = options.maxChange ?? DEFAULT_GOAL_MAX_CHANGE;
	let iterations = 0;
	let best: { x: number; value: number; gap: number } | undefined;
	let negative: Point | undefined;
	let positive: Point | undefined;
	const gapOf = (value: number): number => (Number.isFinite(value) ? Math.abs(value - goal) : NaN);
	const evaluate = (x: number): Point => {
		const value = f(x);
		const gap = gapOf(value);
		const y = value - goal;
		if (!Number.isNaN(gap)) {
			if (!best || gap < best.gap) best = { x, value, gap };
			if (y < 0 && (!negative || Math.abs(y) < Math.abs(negative.y))) negative = { x, y };
			if (y > 0 && (!positive || Math.abs(y) < Math.abs(positive.y))) positive = { x, y };
		}
		return { x, y: Number.isNaN(gap) ? Number.NaN : y };
	};
	const result = (): GoalSolution => ({
		found: !!best && best.gap <= tolerance,
		x: best?.x ?? start,
		value: best?.value ?? Number.NaN,
		iterations,
	});
	const done = (): boolean => !!best && best.gap <= tolerance;

	let previous = evaluate(start);
	if (done()) return result();
	let step = start !== 0 ? Math.abs(start) * 0.1 : 1;
	let current = previous;
	let x = start + step;
	while (iterations < maxIterations) {
		iterations++;
		current = evaluate(x);
		if (done()) break;
		let next = Number.NaN;
		const slopeOk = Number.isFinite(current.y) && Number.isFinite(previous.y);
		if (slopeOk && current.y !== previous.y)
			next = current.x - (current.y * (current.x - previous.x)) / (current.y - previous.y);
		const bracket = negative && positive ? [negative.x, positive.x] : undefined;
		if (bracket) {
			const low = Math.min(...bracket);
			const high = Math.max(...bracket);
			if (!(next > low && next < high)) next = (low + high) / 2;
			if (next === low || next === high) break;
		} else if (!Number.isFinite(next)) {
			// No usable slope and no sign change yet: walk outward with a growing step.
			step *= 2;
			next = x + step;
		}
		if (Number.isFinite(current.y)) previous = current;
		x = next;
	}
	return result();
}
