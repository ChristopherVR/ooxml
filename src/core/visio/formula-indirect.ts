import { parseVisioFormula } from './formula';
import type { VisioFormulaAst, VisioFormulaReference } from './formula';

/** Cells of the shape's own frame that the coordinate conversions read. */
export const FRAME_INPUTS: Record<string, readonly string[]> = {
	ANGLETOLOC: ['Angle', 'FlipX', 'FlipY'],
	LOCTOPAR: ['PinX', 'PinY', 'LocPinX', 'LocPinY', 'Angle', 'FlipX', 'FlipY'],
};

/** The cell an INDIRECT call names when its argument is a literal address such as "Sheet.2!Width". */
export function literalIndirectTarget(node: VisioFormulaAst): VisioFormulaReference | undefined {
	if (node.kind !== 'call' || node.args.length !== 1) return undefined;
	const address = node.args[0]!;
	if (address.kind !== 'string') return undefined;
	try {
		const target = parseVisioFormula(address.value);
		return target.kind === 'reference' &&
			!(target.reference.shapeId === undefined && /^(true|false)$/i.test(target.reference.cell))
			? target.reference
			: undefined;
	} catch {
		return undefined;
	}
}
