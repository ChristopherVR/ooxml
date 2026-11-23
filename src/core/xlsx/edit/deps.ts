// The one place the edit area wires in its sibling modules (number formats and formulas), so a
// signature change there touches a single file here.
export {
	dateToSerial,
	formatValue,
	isDateFormat,
	parseCellInput,
	serialToDate,
} from '../numfmt/index';
export {
	createCalcEngine,
	deleteSheetInFormula,
	formatRefSpec,
	isSpilledCell,
	moveReferencesInFormula,
	renameTableInFormula,
	renameSheetInFormula,
	shiftFormula,
	tokenize,
	translateFormula,
} from '../formula/index';
export { aggregate } from '../formula/functions/math-sum';
export { ErrorSignal, type Scalar } from '../formula/values';
export type {
	CalcEngine,
	CalculationStop,
	CellPosition,
	RefSpec,
	ShiftSpec,
	SpilledCell,
	Token,
} from '../formula/index';
