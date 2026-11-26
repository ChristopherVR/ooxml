import type { AddInIssue, LocalAddIn } from 'ooxml-core/automation/add-ins';
import type { OfficeServiceText } from '../dialog/office-service-dialog';
import { permissionText } from './add-in-local-strings';

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
