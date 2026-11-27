/**
 * `signPackage` as the OPC entry exports it: the signer (its signature builder and the Office
 * signature objects) loads on the first signing, so a consumer that only reads or verifies
 * packages does not download it.
 */
import type { PackageSigner } from './signer';
import type { SignPackageOptions } from './sign-package';

/**
 * `bytes` with one more signature: the host's `signer` signs every part of the package. Fails,
 * with no output, when the package cannot be signed or the result does not verify.
 */
export async function signPackage(
	bytes: Uint8Array,
	signer: PackageSigner,
	options: SignPackageOptions = {},
): Promise<Uint8Array> {
	return (await import('./sign-package')).signPackage(bytes, signer, options);
}
