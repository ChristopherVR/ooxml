import type { PasteOptions } from './types.js';

const MODES = ['all', 'values', 'formats', 'formulas', 'noBorders'] as const;
const OPERATIONS = ['none', 'add', 'subtract', 'multiply', 'divide'] as const;

/** Normalizes command arguments and public paste options without coercing invalid values. */
export function resolvePasteOptions(request: unknown = 'all'): Required<PasteOptions> {
	if (request === 'transpose')
		return { mode: 'all', transpose: true, skipBlanks: false, operation: 'none' };
	if (typeof request === 'string') request = { mode: request };
	if (!request || typeof request !== 'object') throw new RangeError('Invalid paste options.');
	const mode = 'mode' in request ? request.mode : 'all';
	const transpose = 'transpose' in request ? request.transpose : false;
	const skipBlanks = 'skipBlanks' in request ? request.skipBlanks : false;
	const operation = 'operation' in request ? request.operation : 'none';
	if (
		!MODES.some((value) => value === mode) ||
		!OPERATIONS.some((value) => value === operation) ||
		typeof transpose !== 'boolean' ||
		typeof skipBlanks !== 'boolean'
	)
		throw new RangeError('Invalid paste options.');
	return {
		mode: mode as Required<PasteOptions>['mode'],
		transpose,
		skipBlanks,
		operation: operation as Required<PasteOptions>['operation'],
	};
}
