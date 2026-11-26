import {
	AddInManifestError,
	MANIFEST_BYTES,
	PACKAGE_BYTES,
	readLocalAddIn,
	type AddInIssue,
	type LocalAddIn,
	type OfficeAddInHost,
} from 'ooxml-core/automation/add-ins';
import type { OfficeUiDialog } from '../dialog/dialog';
import type { OfficeServiceText } from '../dialog/office-service-dialog';
import type { AddInDialogMode } from './add-in-dialog';
import { localAddInText, permissionText } from './add-in-local-strings';

const installs = new WeakMap<Node, Map<string, LocalAddIn>>();
const LIMIT = 32;

function installationScope(host: HTMLElement | ShadowRoot): Node {
	if (host instanceof ShadowRoot) return host.host;
	// Desktop and mobile Insert controls can be replaced when the active ribbon tab changes.
	const root = host.getRootNode();
	return (
		host.closest('[data-pptx-editor-chrome], [data-pptx-viewer]') ??
		(root instanceof ShadowRoot ? root.host : host)
	);
}

export function mountLocalAddIns(
	dialog: OfficeUiDialog,
	host: HTMLElement | ShadowRoot,
	mode: AddInDialogMode,
	product: OfficeAddInHost,
	locale: string,
	translate: OfficeServiceText,
	{ empty, marketplace }: { empty: HTMLElement | undefined; marketplace: HTMLElement | undefined },
) {
	const doc = dialog.ownerDocument;
	const t = localAddInText(locale, translate);
	const scope = installationScope(host);
	const installed = installs.get(scope) ?? new Map<string, LocalAddIn>();
	installs.set(scope, installed);
	const section = doc.createElement('div');
	dialog.prepend(section);
	const button = (label: string, run: () => void, name?: string) => {
		const control = doc.createElement('button');
		control.type = 'button';
		control.textContent = t(label);
		if (name) control.setAttribute('aria-label', `${t(label)} ${name}`);
		control.addEventListener('click', run);
		return control;
	};
	const appendIssues = (parent: HTMLElement, manifest: LocalAddIn['manifest']) => {
		const format = (issue: AddInIssue) =>
			issue.kind === 'unsupported'
				? `${t('Unsupported feature')} · ${issue.feature}`
				: `${t('Unavailable requirement')} · ${issue.requirement.name} ${issue.requirement.minVersion}`;
		for (const line of [...new Set(manifest.issues.map(format))].slice(0, 12)) {
			const item = doc.createElement('p');
			item.textContent = line;
			parent.append(item);
		}
	};
	if (mode === 'myAddIns') {
		if (!installed.size) return;
		if (empty) empty.hidden = true;
		const list = doc.createElement('ul');
		for (const [id, { manifest }] of installed) {
			const item = doc.createElement('li');
			const title = doc.createElement('p');
			title.textContent = `${manifest.name} (${manifest.version})`;
			const state = doc.createElement('p');
			state.textContent = t('Installed for this editor. Running add-ins is unavailable.');
			item.append(
				title,
				state,
				button(
					'Remove',
					() => {
						installed.delete(id);
						dialog.close();
					},
					manifest.name,
				),
			);
			appendIssues(item, manifest);
			list.append(item);
		}
		section.append(list);
		return;
	}

	const input = doc.createElement('input');
	input.type = 'file';
	input.accept = '.xml,.json,.zip';
	input.setAttribute('aria-label', t('Choose a local add-in file'));
	const caption = doc.createElement('p');
	caption.textContent = t('Choose a local add-in file');
	const preview = doc.createElement('div');
	const status = doc.createElement('p');
	status.setAttribute('role', 'status');
	section.append(caption, input, preview, status);
	let request = 0;
	const alive = (token: number) => request === token && dialog.open && dialog.isConnected;
	input.addEventListener('change', async () => {
		const token = ++request;
		preview.replaceChildren();
		if (marketplace) marketplace.hidden = false;
		status.textContent = '';
		const file = input.files?.[0];
		input.value = '';
		if (!file) return;
		input.disabled = true;
		try {
			const limit = /\.zip$/i.test(file.name) ? PACKAGE_BYTES : MANIFEST_BYTES;
			if (file.size > limit) throw new AddInManifestError('limit');
			const parsed = await readLocalAddIn(await file.arrayBuffer(), { host: product, locale });
			if (!alive(token)) return;
			const { manifest, digest } = parsed;
			const previous = installed.get(manifest.id);
			const heading = doc.createElement('p');
			heading.textContent = `${manifest.name} (${manifest.version})`;
			const summary = doc.createElement('p');
			summary.textContent = manifest.description;
			const details = doc.createElement('dl');
			const detail = (label: string, value: string) => {
				const term = doc.createElement('dt');
				term.textContent = t(label);
				const description = doc.createElement('dd');
				description.textContent = value || t('None');
				details.append(term, description);
			};
			detail('Declared publisher', manifest.publisher);
			detail(
				'Origins this add-in may load',
				[...new Set([...manifest.entryOrigins, ...manifest.declaredOrigins])].join(', '),
			);
			if (previous) detail('Installed version', previous.manifest.version);
			const consent = doc.createElement('section');
			const title = doc.createElement('h3');
			title.textContent = t('Review add-in permissions');
			consent.setAttribute('aria-label', title.textContent);
			const requests = doc.createElement('p');
			requests.textContent = t('This add-in requests:');
			const permissions = doc.createElement('ul');
			for (const permission of manifest.permissions.length ? manifest.permissions : ['']) {
				const item = doc.createElement('li');
				item.textContent = permission ? permissionText(permission, t) : t('None');
				permissions.append(item);
			}
			const notice = doc.createElement('p');
			notice.textContent = t('Install for this editor session. Add-in code stays inactive.');
			consent.append(title, requests, permissions, details, notice);
			preview.append(heading, summary, consent);
			if (marketplace) marketplace.hidden = true;
			appendIssues(consent, manifest);
			if (previous?.digest === digest) {
				status.textContent = t('This add-in is already installed.');
				return;
			}
			const confirm = button(
				previous ? 'Update' : 'Install',
				() => {
					if (!alive(token)) return;
					// A preview may be stale if another dialog installed or removed the same identity.
					if (installed.get(manifest.id) !== previous || (!previous && installed.size >= LIMIT)) {
						status.textContent = t('Could not update the add-in.');
						confirm.disabled = true;
						return;
					}
					installed.set(manifest.id, parsed);
					dialog.close();
				},
				manifest.name,
			);
			consent.append(confirm);
			confirm.focus();
		} catch (error) {
			if (!alive(token)) return;
			const reason =
				error instanceof AddInManifestError
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
			status.textContent = t(reason);
		} finally {
			if (alive(token)) input.disabled = false;
		}
	});
}
