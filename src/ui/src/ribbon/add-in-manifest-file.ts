import {
	AddInManifestError,
	MANIFEST_BYTES,
	PACKAGE_BYTES,
	readLocalAddIn,
	type AddInReadOptions,
} from 'ooxml-core/automation/add-ins';

export { readLocalAddIn } from 'ooxml-core/automation/add-ins';

export async function readAddInInput(
	input: File | Uint8Array | ArrayBuffer,
	options: AddInReadOptions,
) {
	let bytes: Uint8Array;
	if (input instanceof File) {
		const limit = /\.zip$/i.test(input.name) ? PACKAGE_BYTES : MANIFEST_BYTES;
		if (input.size > limit) throw new AddInManifestError('limit');
		bytes = new Uint8Array(await input.arrayBuffer());
	} else {
		if (input.byteLength > PACKAGE_BYTES) throw new AddInManifestError('limit');
		bytes = new Uint8Array(input);
		if (input instanceof ArrayBuffer) bytes = bytes.slice();
	}
	return { bytes, addIn: await readLocalAddIn(bytes, options) };
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
				content: 'Content add-ins are unavailable in this editor.',
				extension: 'This add-in type is unsupported.',
			}[error.code]
		: 'Could not update the add-in.';
}
