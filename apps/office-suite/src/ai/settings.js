/**
 * The assistant's controls: the model menu, adding a model (when the policy lets users), the few
 * personal settings the policy leaves open, and choosing the files a request may read.
 */
import { $, badge, escape, choose, task } from '../ui.js';
import { contextMenu } from '../context-menu.js';
import { icon } from '../icons.js';
import {
	checkEndpoint,
	chosenModel,
	connections,
	currentConnection,
	hasRememberedKey,
	needsKey,
	provider,
	providers,
	removeUserConnection,
	saveUserConnection,
	selectConnection,
	setKey,
} from './providers.js';
import { currentPolicy, isManaged, feature } from './policy.js';
import { personalPreferences, savePersonalPreferences } from './preferences.js';

/** The model menu under the composer's model button. */
export function openModelMenu(event) {
	const policy = currentPolicy();
	const active = currentConnection();
	const items = [];
	for (const c of connections()) {
		const models = c.models?.length > 1 ? c.models : [chosenModel(c)];
		for (const model of models)
			items.push({
				label: c.provider === 'offline' ? c.label : `${c.label} · ${model}`,
				checked: active?.id === c.id && chosenModel(active) === model,
				run: () => {
					selectConnection(c.id, models.length > 1 ? model : undefined);
					const next = currentConnection();
					if (needsKey(next)) askForKey(next);
				},
			});
	}
	if (!items.length) items.push({ label: 'No models are available', disabled: true, run() {} });
	items.push({ separator: true });
	if (policy.allowUserConnections)
		items.push({ label: 'Add a model...', run: () => openConnectionForm() });
	items.push({ label: 'Assistant settings...', run: openPreferences });
	contextMenu(
		event,
		items.map((item) => ({ ...item, error: (e) => console.error(e) })),
	);
}

/** Asks for the key of a connection that needs one before it can be used. */
export function askForKey(connection) {
	const spec = provider(connection.provider);
	const form = document.createElement('form');
	form.innerHTML = `<p><strong>${escape(connection.label)}</strong> needs your key. It is sent only to ${escape(new URL(connection.endpoint).host)}.</p>
<label>${escape(spec.keyLabel ?? 'API key')}<input name="key" type="password" autocomplete="off" required /></label>
<label class="context-option"><input name="remember" type="checkbox" />Remember the key on this device</label>
<button class="primary">Use this key</button>`;
	form.onsubmit = (e) => {
		e.preventDefault();
		const data = new FormData(form);
		setKey(connection.id, String(data.get('key')), data.has('remember'));
		$('dialog').close();
	};
	choose('Connect', form);
}

/** Adds or edits one of the user's own models. */
export function openConnectionForm(existing) {
	const policy = currentPolicy();
	const allowed = providers.filter(
		(p) =>
			p.protocol !== 'offline' && (!policy.userProviders || policy.userProviders.includes(p.id)),
	);
	const form = document.createElement('form');
	form.className = 'ai-settings-form';
	form.innerHTML = `<p>Requests go straight from this page to the provider, with the files you attach. The provider must allow requests from a browser; otherwise ask for an organisation gateway.</p>
<label>Provider<select name="provider">${allowed.map((p) => `<option value="${p.id}" ${p.id === (existing?.provider ?? 'anthropic') ? 'selected' : ''}>${escape(p.name)}</option>`).join('')}</select></label>
<p class="ai-provider-note" data-note></p>
<label>Name<input name="label" maxlength="80" placeholder="Shown in the model menu" /></label>
<label>Endpoint URL<input name="endpoint" type="url" required /></label>
<label>Model<input name="model" required placeholder="Model name" /></label>
<label data-key>API key<input name="key" type="password" autocomplete="off" placeholder="${existing ? 'Leave empty to keep the current key' : 'Kept for this tab unless you choose to remember it'}" /></label>
<label data-key class="context-option"><input name="remember" type="checkbox" ${existing && hasRememberedKey(existing.id) ? 'checked' : ''} />Remember the key on this device</label>
<button class="primary">Connect</button>`;
	const field = (name) => form.elements.namedItem(name);
	const sync = (fresh) => {
		const spec = provider(field('provider').value);
		const same = existing?.provider === spec.id;
		form.querySelector('[data-note]').textContent = spec.note ?? '';
		for (const label of form.querySelectorAll('[data-key]')) label.hidden = spec.auth === 'none';
		if (fresh) {
			field('label').value = (same && existing.label) || spec.name;
			field('endpoint').value = (same && existing.endpoint) || spec.endpoint || '';
			field('model').value = (same && existing.model) || spec.model || '';
		}
	};
	field('provider').onchange = () => sync(true);
	sync(true);
	form.onsubmit = (event) => {
		event.preventDefault();
		task(async () => {
			const values = new FormData(form);
			const spec = provider(values.get('provider'));
			const key = String(values.get('key') ?? '');
			saveUserConnection(
				{
					id: existing?.id,
					label: String(values.get('label')).trim() || spec.name,
					provider: spec.id,
					endpoint: checkEndpoint(values.get('endpoint')),
					model: String(values.get('model')).trim(),
				},
				existing && !key ? undefined : key,
				values.has('remember'),
			);
			$('dialog').close();
		});
	};
	choose(existing ? 'Edit model' : 'Add a model', form);
}

/** The personal settings the policy leaves open, plus the user's own models and keys. */
export function openPreferences() {
	const policy = currentPolicy();
	const controls = policy.userControls;
	const mine = personalPreferences();
	const own = connections().filter((c) => !c.managed);
	const keyed = connections().filter((c) => c.managed && c.keyMode === 'user');
	const form = document.createElement('form');
	form.className = 'ai-settings-form ai-preferences';
	const check = (name, label, on) =>
		`<label class="context-option"><input type="checkbox" name="${name}" ${on ? 'checked' : ''} />${label}</label>`;
	const mark = (c) =>
		`<span class="ai-provider-mark" aria-hidden="true">${escape(provider(c.provider).name[0])}</span>`;
	form.innerHTML = `${isManaged() ? `<p class="ai-managed">${icon('settings')}Your organisation manages the models and most assistant settings.${policy.notice ? ` ${escape(policy.notice)}` : ''}</p>` : ''}
${controls.style || controls.language ? '<h3>Replies</h3>' : ''}
${controls.style ? `<label>Length<select name="style">${['concise', 'balanced', 'detailed'].map((s) => `<option value="${s}" ${(mine.style ?? policy.behaviour.style) === s ? 'selected' : ''}>${s[0].toUpperCase() + s.slice(1)}</option>`).join('')}</select></label>` : ''}
${controls.language ? `<label>Language<input name="language" maxlength="60" placeholder="The language you write in" value="${escape(mine.language ?? '')}" /></label>` : ''}
${controls.temperature ? `<label>Creativity<input name="temperature" type="range" min="0" max="1" step="0.1" value="${mine.temperature ?? policy.behaviour.temperature ?? 0.5}" /></label>` : ''}
${controls.instructions ? `<label>Your instructions<textarea name="instructions" rows="3" maxlength="2000" placeholder="For example: I work in finance; use British spelling.">${escape(mine.instructions ?? '')}</textarea></label>` : ''}
<h3>Behaviour</h3>
${check('attachOpenDocument', 'Attach the open document when the assistant opens beside it', mine.attachOpenDocument !== false)}
${feature('createFiles') ? check('openCreatedFiles', 'Open new files as soon as they are created', mine.openCreatedFiles !== false) : ''}
${check('animate', 'Show replies as they are written', mine.animate !== false)}
${own.length || keyed.length ? '<h3>Models</h3>' : ''}
<div class="ai-models">${[...keyed, ...own]
		.map(
			(c) =>
				`<div><span>${mark(c)}<span><strong>${escape(c.label)}</strong><small>${escape(provider(c.provider).name)} · ${escape(chosenModel(c))}</small></span></span>${c.managed ? `<button type="button" data-key-for="${c.id}">Set key</button>` : `<button type="button" data-edit="${c.id}">Edit</button><button type="button" data-remove="${c.id}">Remove</button>`}</div>`,
		)
		.join('')}</div>
<button class="primary">Save</button>`;
	form.addEventListener('click', (event) => {
		const b = event.target.closest('button[type=button]');
		if (!b) return;
		const c = connections().find(
			(x) => x.id === (b.dataset.edit ?? b.dataset.remove ?? b.dataset.keyFor),
		);
		if (b.dataset.edit) openConnectionForm(c);
		if (b.dataset.keyFor) askForKey(c);
		if (b.dataset.remove) {
			removeUserConnection(c.id);
			b.closest('div').remove();
		}
	});
	form.onsubmit = (event) => {
		event.preventDefault();
		const data = new FormData(form);
		const next = {
			attachOpenDocument: data.has('attachOpenDocument'),
			openCreatedFiles: data.has('openCreatedFiles') || !feature('createFiles'),
			animate: data.has('animate'),
		};
		if (controls.style) next.style = data.get('style');
		if (controls.language) next.language = String(data.get('language')).trim();
		if (controls.temperature) next.temperature = Number(data.get('temperature'));
		if (controls.instructions) next.instructions = String(data.get('instructions')).trim();
		savePersonalPreferences(next);
		$('dialog').close();
	};
	choose('Assistant settings', form);
}

/** Lets the user pick the files (and optionally the Teams channel) sent with the next request. */
export async function chooseContext({ store, selected, includeChat, openIds, onDone }) {
	const docs = (await store.list()).filter((d) => !d.parent);
	const open = docs.filter((d) => openIds.includes(d.id));
	const rest = docs.filter((d) => !openIds.includes(d.id));
	const option = (d) =>
		`<label class="context-option"><input type="checkbox" name="document" value="${d.id}" ${selected.has(d.id) ? 'checked' : ''} />${badge(d.kind)}<span>${escape(d.name)}</span></label>`;
	const form = document.createElement('form');
	form.className = 'ai-context-form';
	form.innerHTML =
		'<p>Only the files you select are sent with your requests in this chat. Type @ in your message to add one quickly.</p>' +
		(docs.length ? '' : '<p>Your library is empty. Create or upload a file first.</p>') +
		(open.length ? `<span class="nav-caption">Open tabs</span>${open.map(option).join('')}` : '') +
		(rest.length ? `<span class="nav-caption">Library</span>${rest.map(option).join('')}` : '') +
		(feature('teamsContext')
			? `<span class="nav-caption">Teams</span><label class="context-option"><input name="chat" type="checkbox" ${includeChat ? 'checked' : ''} />${badge('teams')}<span>Current Teams channel (last 50 messages)</span></label>`
			: '') +
		'<button class="primary">Use selected context</button>';
	form.onsubmit = (event) => {
		event.preventDefault();
		const data = new FormData(form);
		onDone(new Set(data.getAll('document')), data.has('chat'));
		$('dialog').close();
	};
	choose('Attach files', form);
}
