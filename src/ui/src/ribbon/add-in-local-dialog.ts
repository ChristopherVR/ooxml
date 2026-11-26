import {
	AddInManifestError,
	MANIFEST_BYTES,
	PACKAGE_BYTES,
	readLocalAddIn,
	type AddInIssue,
	type LocalAddIn,
	type OfficeAddInHost,
} from 'ooxml-core/automation/add-ins';
import type { OfficeServiceText } from '../dialog/office-service-dialog';
import { permissionText } from './add-in-local-strings';

const installs = new WeakMap<Node, Map<string, LocalAddIn>>();
const LIMIT = 32;

export function localAddIns(scope: Node): Map<string, LocalAddIn> {
	const installed = installs.get(scope) ?? new Map<string, LocalAddIn>();
	installs.set(scope, installed);
	return installed;
}

export async function readAddInFile(file: File, host: OfficeAddInHost, locale: string) {
	const limit = /\.zip$/i.test(file.name) ? PACKAGE_BYTES : MANIFEST_BYTES;
	if (file.size > limit) throw new AddInManifestError('limit');
	return readLocalAddIn(await file.arrayBuffer(), { host, locale });
}

export function installAddIn(
	installed: Map<string, LocalAddIn>,
	preview: LocalAddIn,
	previous: LocalAddIn | undefined,
): boolean {
	const { id } = preview.manifest;
	if (installed.get(id) !== previous || (!previous && installed.size >= LIMIT)) return false;
	installed.set(id, preview);
	return true;
}

export function addInIssueLabels(manifest: LocalAddIn['manifest'], t: OfficeServiceText): string[] {
	const format = (issue: AddInIssue) =>
		issue.kind === 'unsupported'
			? `${t('Unsupported feature')} · ${issue.feature}`
			: `${t('Unavailable requirement')} · ${issue.requirement.name} ${issue.requirement.minVersion}`;
	return [...new Set(manifest.issues.map(format))].slice(0, 12);
}

export function addInConsent(
	preview: LocalAddIn,
	previous: LocalAddIn | undefined,
	t: OfficeServiceText,
) {
	const { manifest } = preview;
	const characters = Array.from(manifest.description);
	const oldVersion = previous?.manifest.version;
	const newVersion = manifest.version;
	return {
		description:
			characters.length > 250 ? `${characters.slice(0, 249).join('')}…` : manifest.description,
		permissions: manifest.permissions.length
			? manifest.permissions.map((permission) => permissionText(permission, t))
			: [t('None')],
		origins: [...new Set([...manifest.entryOrigins, ...manifest.declaredOrigins])].join(', '),
		update:
			oldVersion && previous?.digest !== preview.digest
				? oldVersion === newVersion
					? t('Update this add-in and keep version {version}').replace('{version}', newVersion)
					: t('Update from {oldVersion} to {newVersion}')
							.replace('{oldVersion}', oldVersion)
							.replace('{newVersion}', newVersion)
				: undefined,
	};
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
