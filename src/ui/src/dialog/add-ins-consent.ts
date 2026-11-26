import { html, nothing } from 'lit';
import { glyph } from '../glyph';
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
					[
						...new Set(
							manifest.requirements.map(
								({ name, minVersion, maxVersion }) =>
									`${name} ${minVersion}${maxVersion ? ` - ${maxVersion}` : ''}`,
							),
						),
					].join(', ') || t('None')
				}</dd
			>
			<dt>${t('Icons')}</dt>
			<dd
				>${
					icons.length
						? icons.map((icon) =>
								icon.kind === 'package' && icon.preview
									? html`<img
											src=${icon.preview}
											width="32"
											height="32"
											alt=${t('Icon preview')}
										/>`
									: html`<span class="icon-placeholder">
											${glyph('addIns', 'tile-icon')}
											${icon.kind === 'remote' ? new URL(icon.url).hostname : t('Icon preview unavailable')}
										</span>`,
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
	options: { persistent: boolean; reason: string | undefined },
) {
	const { manifest } = addIn;
	const { description, update } = addInConsent(addIn, previous, t);
	const already = previous?.digest === addIn.digest;
	return html`<section aria-labelledby="consent-title" class="consent">
		<h3 id="consent-title" tabindex="-1">${t('Review add-in permissions')}</h3>
		<p class="tile-name">${manifest.name} (${manifest.version})</p><p>${description}</p>
		${update ? html`<p>${update}</p>` : nothing} ${renderAddInDetails(addIn, t)}
		<p
			>${t(options.persistent ? 'Save this add-in in host preferences.' : 'Install for this editor session.')}</p
		>
		<p>${options.reason ?? t('This add-in runs only when opened.')}</p>
		${addInIssueLabels(manifest, t).map((line) => html`<p>${line}</p>`)}
		${already ? html`<p>${t('This add-in is already installed.')}</p>` : nothing}
	</section>`;
}

export function renderAddInConsentActions(
	addIn: LocalAddIn,
	previous: LocalAddIn | undefined,
	t: OfficeServiceText,
	actions: { confirm: () => void; cancel: () => void; busy: boolean },
) {
	const label = t(previous ? 'Update' : 'Install');
	const already = previous?.digest === addIn.digest;
	return html`<button type="button" slot="footer" ?disabled=${actions.busy} @click=${actions.cancel}
			>${t('Cancel')}</button
		><button
			type="button"
			slot="footer"
			class="primary"
			aria-label=${`${label} ${addIn.manifest.name}`}
			?disabled=${already || actions.busy}
			@click=${actions.confirm}
			>${label}</button
		>`;
}
