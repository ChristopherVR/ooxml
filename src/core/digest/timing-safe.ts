/**
 * Constant-time byte comparison for password verifiers, MACs and digests. A length mismatch
 * fails at once (lengths are not secret); equal-length inputs are always scanned in full, so
 * the running time does not reveal how many leading bytes matched.
 */
export function timingSafeEqual(a: Uint8Array, b: Uint8Array): boolean {
	if (a.length !== b.length) return false;
	let difference = 0;
	for (let i = 0; i < a.length; i++) difference |= a[i]! ^ b[i]!;
	return difference === 0;
}
