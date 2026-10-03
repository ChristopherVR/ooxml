// Compatibility entry point. The implementation moved to the shared crypto area (src/crypto).
export {
	decryptAgilePackage,
	decryptStandardPackage,
	verifyAgileDataIntegrity,
	verifyAgilePassword,
	verifyStandardPassword,
} from '../../../crypto/decrypt.js';
