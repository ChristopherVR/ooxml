import { sha256 } from '../../digest/sha256';
import { decodeXmlBytes, xmlEncodingOf } from '../../xml/decode';
import type { AddInReadOptions, LocalAddIn } from './model';
import { decodeManifest, openManifestPackage } from './package';
import { readUnifiedManifest } from './unified';
import { AddInManifestError, MANIFEST_BYTES, PACKAGE_BYTES } from './validation';
import { readXmlManifest } from './xml';

export type {
	AddInCommand,
	AddInIcon,
	AddInIssue,
	AddInManifest,
	AddInReadOptions,
	AddInRequirement,
	AddInTab,
	LocalAddIn,
	OfficeAddInHost,
} from './model';
export { AddInManifestError, MANIFEST_BYTES, PACKAGE_BYTES } from './validation';

/** Reads a detached local manifest or a unified ZIP. URLs, schemas and assets never trigger I/O. */
export async function readLocalAddIn(
	input: Uint8Array | ArrayBuffer,
	options: AddInReadOptions,
): Promise<LocalAddIn> {
	const bytes = input instanceof Uint8Array ? input : new Uint8Array(input);
	if (!bytes.length || bytes.length > PACKAGE_BYTES)
		throw new AddInManifestError('limit', 'input size');
	const owned = new Uint8Array(bytes);
	const zipped = owned[0] === 0x50 && owned[1] === 0x4b;
	const digest = Array.from(sha256(owned), (byte) => byte.toString(16).padStart(2, '0')).join('');
	if (zipped) {
		const assets = await openManifestPackage(owned);
		return {
			digest,
			manifest: await readUnifiedManifest(await assets.readText('manifest.json'), assets, options),
		};
	}
	if (owned.byteLength > MANIFEST_BYTES) throw new AddInManifestError('limit', 'manifest size');
	if (xmlEncodingOf(owned) !== 'utf-8' || decodeManifest(owned).trimStart().startsWith('<')) {
		let source: string;
		try {
			source = decodeXmlBytes(owned);
		} catch {
			throw new AddInManifestError('manifest', 'XML encoding');
		}
		return { digest, manifest: readXmlManifest(source, options) };
	}
	const assets = {
		has: () => false,
		readText: (): never => {
			throw new AddInManifestError('package', 'missing locale file');
		},
	};
	return { digest, manifest: await readUnifiedManifest(decodeManifest(owned), assets, options) };
}
