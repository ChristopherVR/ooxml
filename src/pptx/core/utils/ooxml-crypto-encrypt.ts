// Compatibility entry point. The implementation moved to the shared crypto area (src/crypto).
import { encryptOoxmlPackage } from '../../../crypto/index.js';
import type { EncryptionOptions } from '../../../crypto/types.js';

export { encryptAgilePackage, encryptStandardPackage } from '../../../crypto/encrypt.js';
export {
	buildAgileEncryptionInfoXml,
	buildEncryptionInfoStream,
	buildStandardEncryptionInfoStream,
} from '../../../crypto/encryption-info-write.js';

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
