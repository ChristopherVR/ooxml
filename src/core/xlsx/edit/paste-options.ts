import type { PasteOptions } from './types.js';

const MODES = ['all', 'values', 'formats', 'formulas'] as const;

/** Normalizes command arguments and public paste options without coercing invalid values. */
export function resolvePasteOptions(request: unknown = 'all'): Required<PasteOptions> {
	if (request === 'transpose') return { mode: 'all', transpose: true, skipBlanks: false };
	if (typeof request === 'string') request = { mode: request };
	if (!request || typeof request !== 'object') throw new RangeError('Invalid paste options.');
	const mode = 'mode' in request ? request.mode : 'all';
	const transpose = 'transpose' in request ? request.transpose : false;
	const skipBlanks = 'skipBlanks' in request ? request.skipBlanks : false;
	if (
		!MODES.some((value) => value === mode) ||
		typeof transpose !== 'boolean' ||
		typeof skipBlanks !== 'boolean'
	)
		throw new RangeError('Invalid paste options.');
	return { mode: mode as Required<PasteOptions>['mode'], transpose, skipBlanks };
}
