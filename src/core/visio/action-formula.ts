import { executableCellFormula } from './cell-formula';
import { analyzeVisioFormula, parseVisioFormula, type VisioFormulaAst } from './formula';
import type { VisioActionRun } from './shape-sections-types';

/** One `SETF(GetRef(cell), value)` step of an action. */
export interface VisioSetfStep {
	/** The cell name as the formula wrote it (`Geometry2.NoShow`, `User.Mode`). */
	target: string;
	value: VisioFormulaAst;
}

const unsupported = (reason: string): { reason: string } => ({ reason });

/**
 * The SETF steps of an Action formula, or why it is not one this editor runs. An action runs when
 * it is `SETF(GetRef(cell), value)` on a cell of the shape itself, or several of them added
 * together, whose value is computed from numbers and cells (`NOT(Geometry2.NoShow)`).
 */
export function parseActionSteps(
	source: string | undefined,
): { steps: VisioSetfStep[] } | { reason: string } {
	const formula = executableCellFormula(source);
	if (!formula) return unsupported('The action has no formula.');
	let ast: VisioFormulaAst;
	try {
		ast = parseVisioFormula(formula);
	} catch {
		return unsupported('The action formula is not one this editor can read.');
	}
	const steps: VisioSetfStep[] = [];
	const visit = (node: VisioFormulaAst): string | undefined => {
		if (node.kind === 'binary' && node.operator === '+')
			return visit(node.left) ?? visit(node.right);
		if (node.kind !== 'call') return 'The action is not a call to SETF.';
		if (node.name !== 'SETF')
			return `${node.name} runs outside the ShapeSheet and is not run here.`;
		const [reference, value] = node.args;
		if (node.args.length !== 2 || !reference || !value) return 'SETF needs a target and a value.';
		if (reference.kind !== 'call' || reference.name !== 'GETREF' || reference.args.length !== 1)
			return 'SETF must name its target with GetRef(cell).';
		const target = reference.args[0];
		if (target?.kind !== 'reference' || target.reference.shapeId !== undefined)
			return 'SETF can only change a cell of the shape that owns the action.';
		try {
			const analysis = analyzeVisioFormula(value);
			if (analysis.dynamic || analysis.unsupportedFunctions.length)
				return 'The value of SETF uses functions this editor cannot evaluate.';
		} catch {
			return 'The value of SETF is not one this editor can read.';
		}
		steps.push({ target: target.reference.cell, value });
		return undefined;
	};
	const reason = visit(ast);
	return reason ? { reason } : steps.length ? { steps } : unsupported('The action has no steps.');
}

/** What the model records about running an action. */
export function actionRun(source: string | undefined): VisioActionRun {
	const parsed = parseActionSteps(source);
	return 'steps' in parsed
		? { kind: 'setf', targets: parsed.steps.map((step) => step.target) }
		: { kind: 'unsupported', reason: parsed.reason };
}
