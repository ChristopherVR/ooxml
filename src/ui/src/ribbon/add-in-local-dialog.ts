import type { AddInIssue, LocalAddIn } from 'ooxml-core/automation/add-ins';
import type { OfficeServiceText } from '../dialog/office-service-dialog';
import { permissionText } from './add-in-local-strings';

const featureLabels = new Map([
	['executefunction', 'Function commands'],
	['built-in ribbon tab', 'Built-in ribbon tab placement'],
	['automatic task pane', 'Automatic task panes'],
	['web application identity', 'Microsoft sign-in'],
	['contextmenu', 'Context menus'],
	['togglebutton', 'Toggle commands'],
	['shared runtime', 'Shared runtimes'],
	['mobile form factor', 'Mobile add-ins'],
	['dynamic ribbon command', 'Dynamic ribbon commands'],
	['versionoverrides', 'Manifest version overrides'],
]);

export function addInIssueLabel(issue: AddInIssue, t: OfficeServiceText): string {
	if (issue.kind === 'requirement')
		return `${t('Unavailable requirement')} · ${issue.requirement.name} ${issue.requirement.minVersion}`;
	const feature = featureLabels.get(issue.feature.toLowerCase());
	return feature ? `${t('Unsupported feature')} · ${t(feature)}` : t('Unsupported feature');
}

export function addInIssueLabels(manifest: LocalAddIn['manifest'], t: OfficeServiceText): string[] {
	return [...new Set(manifest.issues.map((issue) => addInIssueLabel(issue, t)))].slice(0, 12);
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
