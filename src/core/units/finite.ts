/**
 * `Number.isFinite` as a type guard: true only for a finite number, so a caller holding
 * `number | undefined` (an indexed read under `noUncheckedIndexedAccess`) narrows to `number`.
 */
export function isFiniteNumber(value: unknown): value is number {
	return Number.isFinite(value);
}
