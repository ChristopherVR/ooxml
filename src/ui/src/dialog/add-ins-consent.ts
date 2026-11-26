import { html, nothing } from 'lit';
import type { LocalAddIn } from 'ooxml-core/automation/add-ins';
import { addInConsent, addInIssueLabels } from '../ribbon/add-in-local-dialog';
import type { OfficeServiceText } from './office-service-dialog';

const hostNames = { word: 'Word', excel: 'Excel', powerpoint: 'PowerPoint' };

export function renderAddInDetails(addIn: LocalAddIn, t: OfficeServiceText) {
	const { manifest } = addIn;
	const { permissions, origins } = addInConsent(addIn, undefined, t);
	const icons = manifest.icons.slice(0, 8);
	return html`<dl>
			<dt>${t('Declared publisher')}</dt><dd>${manifest.publisher}</dd>
			<dt>${t('Supported hosts')}</dt
			><dd>${manifest.hosts.map((host) => hostNames[host]).join(', ')}</dd>
			<dt>${t('Requirements')}</dt>
			<dd
				>${
					manifest.requirements
						.map(
							({ name, minVersion, maxVersion }) =>
								`${name} ${minVersion}${maxVersion ? ` - ${maxVersion}` : ''}`,
						)
						.join(', ') || t('None')
				}</dd
			>
			<dt>${t('Icons')}</dt>
			<dd
				>${
					icons.length
						? icons.map((icon) =>
								icon.kind === 'package' && icon.preview
									? html`<img src=${icon.preview} width="32" height="32" alt=${icon.path} />`
									: html`<div>${icon.kind === 'package' ? icon.path : icon.url}</div>`,
							)
						: t('None')
				}</dd
			>
			<dt>${t('Origins this add-in may load')}</dt><dd>${origins || t('None')}</dd>
		</dl>
		<p>${t('This add-in requests:')}</p>
		<ul
			>${permissions.map((permission) => html`<li>${permission}</li>`)}</ul
		>`;
}

export function renderAddInConsent(
	addIn: LocalAddIn,
	previous: LocalAddIn | undefined,
	t: OfficeServiceText,
	actions: { confirm: () => void; cancel: () => void },
) {
	const { manifest } = addIn;
	const { description, update } = addInConsent(addIn, previous, t);
	const already = previous?.digest === addIn.digest;
	const label = t(previous ? 'Update' : 'Install');
	return html`<section aria-labelledby="consent-title" class="consent">
		<h3 id="consent-title" tabindex="-1">${t('Review add-in permissions')}</h3>
		<p class="tile-name">${manifest.name} (${manifest.version})</p><p>${description}</p>
		${update ? html`<p>${update}</p>` : nothing} ${renderAddInDetails(addIn, t)}
		<p>${t('Install for this editor session. Add-in code stays inactive.')}</p>
		${addInIssueLabels(manifest, t).map((line) => html`<p>${line}</p>`)}
		${already ? html`<p>${t('This add-in is already installed.')}</p>` : nothing}
		<div class="consent-actions">
			<button
				type="button"
				class="primary"
				aria-label=${`${label} ${manifest.name}`}
				?disabled=${already}
				@click=${actions.confirm}
				>${label}</button
			>
			<button type="button" @click=${actions.cancel}>${t('Cancel')}</button>
		</div>
	</section>`;
}
