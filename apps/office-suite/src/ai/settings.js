/** The assistant's two dialogs: which model to talk to, and which files it may read. */
import { $, badge, escape, choose, task } from '../ui.js';
import {
	checkEndpoint,
	currentConnection,
	provider,
	providers,
	saveConnection,
} from './providers.js';

export function openSettings() {
	const current = currentConnection();
	const form = document.createElement('form');
	form.className = 'ai-settings-form';
	form.innerHTML = `<p>Choose where the assistant sends your requests. Attached file content goes to that provider and nowhere else; the offline preview sends nothing.</p>
<label>Provider<select name="provider">${providers.map((p) => `<option value="${p.id}" ${p.id === current.provider ? 'selected' : ''}>${escape(p.name)}</option>`).join('')}</select></label>
<p class="ai-provider-note" data-note></p>
<label data-remote>Endpoint URL<input name="endpoint" type="url" placeholder="https://your-provider/v1/chat/completions" /></label>
<label data-remote>Model<input name="model" placeholder="Model name" /></label>
<label data-key>API key<input name="key" type="password" autocomplete="off" placeholder="Kept for this tab unless you choose to remember it" /></label>
<label data-key class="context-option"><input name="remember" type="checkbox" />Remember the key on this device</label>
<button class="primary">Connect</button>`;
	const field = (name) => form.elements.namedItem(name);
	const sync = (fresh) => {
		const spec = provider(field('provider').value);
		const same = spec.id === current.provider;
		for (const label of form.querySelectorAll('[data-remote]'))
			label.hidden = spec.protocol === 'offline';
		for (const label of form.querySelectorAll('[data-key]'))
			label.hidden = spec.protocol === 'offline' || !spec.key;
		field('endpoint').required = field('model').required = spec.protocol !== 'offline';
		form.querySelector('[data-note]').textContent = spec.note ?? '';
		if (fresh) {
			field('endpoint').value = (same && current.endpoint) || spec.endpoint || '';
			field('model').value =
				(same && current.model !== 'Offline preview' && current.model) || spec.model || '';
			field('key').value = (same && current.key) || '';
			field('remember').checked = same && Boolean(current.remember);
		}
	};
	field('provider').onchange = () => sync(true);
	sync(true);
	form.onsubmit = (event) => {
		event.preventDefault();
		task(async () => {
			const values = new FormData(form);
			const spec = provider(values.get('provider'));
			if (spec.protocol === 'offline')
				saveConnection({ provider: 'offline', model: 'Offline preview' }, false);
			else {
				const remember = values.has('remember');
				saveConnection(
					{
						provider: spec.id,
						endpoint: checkEndpoint(values.get('endpoint')),
						model: String(values.get('model')).trim(),
						key: String(values.get('key') ?? ''),
						remember,
					},
					remember,
				);
			}
			$('dialog').close();
		});
	};
	choose('Assistant model', form);
}

/** Lets the user pick the files (and optionally the Teams channel) sent with the next request. */
export async function chooseContext({ store, selected, includeChat, onDone }) {
	const docs = (await store.list()).filter((d) => !d.parent);
	const form = document.createElement('form');
	form.className = 'ai-context-form';
	form.innerHTML =
		'<p>Only the files you select, and the Teams channel if you include it, are sent with your requests in this chat.</p>' +
		(docs.length ? '' : '<p>Your library is empty. Create or upload a file first.</p>') +
		docs
			.map(
				(d) =>
					`<label class="context-option"><input type="checkbox" name="document" value="${d.id}" ${selected.has(d.id) ? 'checked' : ''} />${badge(d.kind)}<span>${escape(d.name)}</span></label>`,
			)
			.join('') +
		`<label class="context-option"><input name="chat" type="checkbox" ${includeChat ? 'checked' : ''} />${badge('teams')}<span>Current Teams channel (last 50 messages)</span></label><button class="primary">Use selected context</button>`;
	form.onsubmit = (event) => {
		event.preventDefault();
		const data = new FormData(form);
		onDone(new Set(data.getAll('document')), data.has('chat'));
		$('dialog').close();
	};
	choose('Attach files', form);
}
