import type { ModernPasswordHash } from '../model.js';
import { modernHashAttributeNames } from '../read/password-hash.js';

/**
 * The attribute values of an agile hash, keyed by their XML names (`algorithmName`, or
 * `workbookAlgorithmName` with the `workbook` prefix). Written only from the model, never from
 * the source XML: a hash the model dropped (password removed or changed) must not come back.
 */
export function modernHashValues(
	hash: ModernPasswordHash | undefined,
	prefix = '',
): Record<string, string | undefined> {
	if (!hash) return {};
	const names = modernHashAttributeNames(prefix);
	return {
		[names.algorithmName]: hash.algorithmName,
		[names.hashValue]: hash.hashValue,
		[names.saltValue]: hash.saltValue,
		[names.spinCount]: hash.spinCount === undefined ? undefined : String(hash.spinCount),
	};
}
