// Compatibility entry point. The implementation moved to the shared crypto area (src/crypto).
export {
	BLOCK_KEYS,
	computeStandardKeyBase,
	deriveAgileKey,
	deriveStandardKey,
	deriveStandardKeyFromBase,
	generateIV,
} from '../../../crypto/key-derivation.js';
export { RC4_ALG_ID, parseEncryptionInfo } from '../../../crypto/encryption-info.js';
