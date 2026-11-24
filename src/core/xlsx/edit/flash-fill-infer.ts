import {
	type AtomExpr,
	type FlashFillExample,
	evalAtom,
	generateAtomsForExamples,
} from './flash-fill-atoms';

export type { FlashFillExample } from './flash-fill-atoms';

type ProgramStep = { kind: 'literal'; text: string } | { kind: 'atom'; expr: AtomExpr };

export type FlashFillProgram = (inputs: readonly string[]) => string | undefined;

interface SearchNode {
	positions: number[];
	cost: number;
	hasAtom: boolean;
	steps: ProgramStep[];
}

function isSpuriousSingleExample(steps: readonly ProgramStep[]): boolean {
	let atomChars = 0;
	let literalChars = 0;
	let hasAlphaLiteral = false;
	let onlyTinySlice = true;
	for (const step of steps) {
		if (step.kind === 'literal') {
			literalChars += step.text.length;
			if (/[A-Za-z0-9]/.test(step.text)) hasAlphaLiteral = true;
		} else {
			const k = step.expr.op.kind;
			if (k !== 'prefix' && k !== 'suffix' && k !== 'slice') onlyTinySlice = false;
			if (k === 'prefix' || k === 'suffix' || k === 'slice') atomChars += step.expr.op.len;
			else atomChars += 2;
		}
	}
	return onlyTinySlice && atomChars <= 1 && literalChars >= 2 && hasAlphaLiteral;
}

/**
 * Infers the simplest string transformation program matching every `(inputs, output)` example,
 * or returns `undefined` when no consistent pattern exists.
 */
export function inferFlashFillProgram(
	examples: readonly FlashFillExample[],
): FlashFillProgram | undefined {
	const valid = examples.filter((ex) => ex.output.length > 0);
	if (!valid.length) return undefined;
	const first = valid[0]!;
	const atoms = generateAtomsForExamples(first, valid);
	if (!atoms.length) return undefined;

	const evalTable = atoms.map((expr) => valid.map((ex) => evalAtom(expr, ex.inputs) ?? ''));
	const targets = valid.map((ex) => ex.output);
	const goalKey = targets.map((t) => t.length).join(',');

	const queue: SearchNode[] = [
		{ positions: targets.map(() => 0), cost: 0, hasAtom: false, steps: [] },
	];
	const best = new Map<string, number>();
	let expanded = 0;
	const MAX_EXPANDED = 2500;

	while (queue.length && expanded < MAX_EXPANDED) {
		let bestIdx = 0;
		for (let i = 1; i < queue.length; i++) if (queue[i]!.cost < queue[bestIdx]!.cost) bestIdx = i;
		const cur = queue.splice(bestIdx, 1)[0]!;
		const key = `${cur.positions.join(',')}|${cur.hasAtom ? 1 : 0}`;
		if ((best.get(key) ?? Infinity) <= cur.cost) continue;
		best.set(key, cur.cost);
		expanded++;

		if (cur.positions.join(',') === goalKey) {
			if (!cur.hasAtom) continue;
			if (valid.length === 1 && isSpuriousSingleExample(cur.steps)) continue;
			const finalSteps = cur.steps;
			return (inputs: readonly string[]): string | undefined => {
				let out = '';
				for (const step of finalSteps) {
					if (step.kind === 'literal') out += step.text;
					else {
						const piece = evalAtom(step.expr, inputs);
						if (piece === undefined) return undefined;
						out += piece;
					}
				}
				return out;
			};
		}

		for (let a = 0; a < atoms.length; a++) {
			const vals = evalTable[a]!;
			const nextPos: number[] = [];
			let ok = true;
			for (let m = 0; m < targets.length; m++) {
				const v = vals[m]!;
				const p = cur.positions[m]!;
				if (!v || !targets[m]!.startsWith(v, p)) {
					ok = false;
					break;
				}
				nextPos.push(p + v.length);
			}
			if (!ok) continue;
			const expr = atoms[a]!;
			queue.push({
				positions: nextPos,
				cost: cur.cost + expr.cost,
				hasAtom: true,
				steps: [...cur.steps, { kind: 'atom', expr }],
			});
		}

		const ch = targets[0]![cur.positions[0]!];
		if (ch !== undefined) {
			let shared = true;
			const nextPos: number[] = [];
			for (let m = 0; m < targets.length; m++) {
				const p = cur.positions[m]!;
				if (targets[m]![p] !== ch) {
					shared = false;
					break;
				}
				nextPos.push(p + 1);
			}
			if (shared) {
				const prev = cur.steps[cur.steps.length - 1];
				const isSep = /[\s,.\-_/\\@:;|()[\]]/.test(ch);
				const charCost = isSep ? 0.3 : 3.0;
				const steps: ProgramStep[] =
					prev?.kind === 'literal'
						? [...cur.steps.slice(0, -1), { kind: 'literal', text: prev.text + ch }]
						: [...cur.steps, { kind: 'literal', text: ch }];
				const openCost = prev?.kind === 'literal' ? charCost : (isSep ? 1.0 : 5.0) + charCost;
				queue.push({
					positions: nextPos,
					cost: cur.cost + openCost,
					hasAtom: cur.hasAtom,
					steps,
				});
			}
		}
	}

	return undefined;
}
