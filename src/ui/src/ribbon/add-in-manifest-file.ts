import {
	AddInManifestError,
	MANIFEST_BYTES,
	PACKAGE_BYTES,
	readLocalAddIn,
	type OfficeAddInHost,
} from 'ooxml-core/automation/add-ins';

export async function readAddInFile(file: File, host: OfficeAddInHost, locale: string) {
	const limit = /\.zip$/i.test(file.name) ? PACKAGE_BYTES : MANIFEST_BYTES;
	if (file.size > limit) throw new AddInManifestError('limit');
	return readLocalAddIn(await file.arrayBuffer(), { host, locale });
}

export function addInError(error: unknown): string {
	return error instanceof AddInManifestError
		? {
				manifest: 'The add-in manifest is invalid.',
				package: 'The add-in package is invalid or missing a file.',
				limit: 'The add-in file exceeds the size or entry limit.',
				url: 'The add-in specifies an unsafe URL or origin.',
				host: 'The add-in does not support this editor.',
				version: 'This manifest version is unsupported.',
				extension: 'This add-in type is unsupported.',
			}[error.code]
		: 'Could not update the add-in.';
}
