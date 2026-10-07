// Compatibility entry point. The implementation moved to the shared crypto area (src/core/crypto).
import { encryptOoxmlPackage } from '../../../crypto/index';
import type { EncryptionOptions } from '../../../crypto/types';

export { encryptAgilePackage, encryptStandardPackage } from '../../../crypto/encrypt';
export {
	buildAgileEncryptionInfoXml,
	buildEncryptionInfoStream,
	buildStandardEncryptionInfoStream,
} from '../../../crypto/encryption-info-write';

/** Former name of `encryptOoxmlPackage(bytes, password, { encryptionScheme: 'standard' })`. */
export async function encryptPptxStandard(
	pptxBuffer: ArrayBuffer,
	password: string,
	options?: EncryptionOptions,
): Promise<ArrayBuffer> {
	const bytes = await encryptOoxmlPackage(pptxBuffer, password, {
		...options,
		encryptionScheme: 'standard',
	});
	return bytes.buffer as ArrayBuffer;
}
