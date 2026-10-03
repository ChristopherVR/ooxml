// Compatibility entry point. The implementation moved to the shared crypto area (src/crypto).
export {
	EncryptedFileError,
	detectFileFormat,
	type FileFormatDetection,
} from '../../../crypto/detect.js';
