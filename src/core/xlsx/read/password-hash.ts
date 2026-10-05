import type { XmlElement } from '../../xml/index.js';
import type { ModernPasswordHash } from '../model.js';
import { att } from './xml-util.js';

/** The attribute names of an agile hash: `algorithmName`, or `workbookAlgorithmName` with a prefix. */
export function modernHashAttributeNames(prefix = '') {
	const name = (base: string) =>
		prefix ? `${prefix}${base[0]!.toUpperCase()}${base.slice(1)}` : base;
	return {
		algorithmName: name('algorithmName'),
		hashValue: name('hashValue'),
		saltValue: name('saltValue'),
		spinCount: name('spinCount'),
	};
}

/**
 * Reads an ECMA-376 agile password hash from `node` (`sheetProtection`, or `workbookProtection`
 * with the `workbook` prefix). A missing or non-numeric `spinCount` is left out of the model.
 */
export function readModernHash(
	node: XmlElement | undefined,
	prefix = '',
): ModernPasswordHash | undefined {
	const names = modernHashAttributeNames(prefix);
	const algorithmName = att(node, names.algorithmName);
	const hashValue = att(node, names.hashValue);
	const saltValue = att(node, names.saltValue);
	if (!algorithmName || !hashValue || !saltValue) return undefined;
	const hash: ModernPasswordHash = { algorithmName, hashValue, saltValue };
	const spin = att(node, names.spinCount);
	const spinCount = spin === undefined || spin.trim() === '' ? Number.NaN : Number(spin);
	if (Number.isFinite(spinCount)) hash.spinCount = spinCount;
	return hash;
}
