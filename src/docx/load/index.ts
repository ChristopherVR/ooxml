// Document format detection and loading: DOCX (OPC zip), password-protected (encrypted) .docx
// through the shared crypto area, and Word 97-2003 .doc (OLE2 compound file). The legacy reader
// and the encrypted-package container inline the shared `@christophervr/ole2` codecs at build
// time (see tsup.config.ts).
export {
	detectDocumentFormat,
	loadDocument,
	type DocumentFormat,
	type LoadDocumentOptions,
} from './detect.js';
export { LegacyDocError, loadLegacyDoc } from './legacy-doc.js';
export {
	DataIntegrityError,
	IncorrectPasswordError,
	PasswordRequiredError,
	isOoxmlCryptoError,
	type OoxmlCryptoErrorCode,
} from '../../crypto/errors.js';
export {
	encryptOoxmlPackage,
	isEncryptedOoxmlPackage,
	type EncryptionOptions,
} from '../../crypto/index.js';
