import { $, escape, choose, notify, task } from './ui.js';
import { inspect, operationHelp, prepareChanges, validateChanges } from './ai-operations.js';

export function mountAssistant({ store, sessions, refresh, getActive, getTeams }) {
	let connection = null,
		selected = new Set(),
		includeChat = false,
		busy = false;
	function message(text, user = false) {
		const node = document.createElement('div');
		node.className = `ai-message${user ? ' user' : ''}`;
		node.textContent = text;
		$('ai-messages').append(node);
		node.scrollIntoView({ block: 'nearest' });
		return node;
	}
	$('ai-settings').onclick = () => {
		const form = document.createElement('form');
		form.innerHTML = `<p>Connect an OpenAI-compatible chat completions endpoint. Selected file content is sent to this provider. Connection credentials are kept only for this session.</p><label>Endpoint URL<input name="endpoint" type="url" placeholder="https://your-provider/v1/chat/completions" value="${escape(connection?.endpoint ?? '')}" required /></label><label>Model<input name="model" value="${escape(connection?.model ?? '')}" required /></label><label>API key<input name="key" type="password" autocomplete="off" /></label><button class="primary">Connect</button>`;
		form.onsubmit = (event) => {
			event.preventDefault();
			task(async () => {
				const values = new FormData(form),
					url = new URL(values.get('endpoint'));
				if (
					url.username ||
					url.password ||
					url.hash ||
					(url.protocol !== 'https:' &&
						!(
							url.protocol === 'http:' && ['localhost', '127.0.0.1', '[::1]'].includes(url.hostname)
						))
				)
					throw new Error('Use HTTPS, or a local endpoint on localhost.');
				connection = { endpoint: url.href, model: values.get('model'), key: values.get('key') };
				$('ai-state').textContent = `Connected: ${connection.model}`;
				$('dialog').close();
			});
		};
		choose('Assistant connection', form);
	};
	$('ai-context').onclick = () =>
		task(async () => {
			const docs = await store.list();
			if (!selected.size && getActive()) selected.add(getActive());
			const form = document.createElement('form');
			form.innerHTML =
				'<p>Only selected files and the optional current Teams conversation are sent with your next request.</p>' +
				docs
					.map(
						(d) =>
							`<label class="context-option"><input type="checkbox" name="document" value="${d.id}" ${selected.has(d.id) ? 'checked' : ''} />${escape(d.name)}</label>`,
					)
					.join('') +
				`<label class="context-option"><input name="chat" type="checkbox" ${includeChat ? 'checked' : ''} />Current Teams channel (last 50 messages)</label><button class="primary">Use selected context</button>`;
			form.onsubmit = (event) => {
				event.preventDefault();
				const data = new FormData(form);
				selected = new Set(data.getAll('document'));
				includeChat = data.has('chat');
				$('ai-scope').textContent =
					`${selected.size} file${selected.size === 1 ? '' : 's'} selected${includeChat ? ' + Teams channel' : ''}`;
				$('dialog').close();
			};
			choose('Assistant context', form);
		});
	async function verify(docs) {
		for (const doc of docs) {
			sessions.sessions.get(doc.id)?.editor?.flush?.();
			if (sessions.isDirty(doc.id) || (await store.get(doc.id)).revision !== doc.revision)
				throw new Error(`${doc.name} changed. Ask again using the latest version.`);
		}
	}
	function proposal(changes, docs) {
		validateChanges(changes, docs);
		const panel = document.createElement('div');
		panel.className = 'ai-proposal';
		const description = document.createElement('p');
		description.textContent = `${changes.length} proposed edit${changes.length === 1 ? '' : 's'}. Review the exact changes below.`;
		const preview = document.createElement('pre');
		preview.textContent = JSON.stringify(
			changes.map((c) => ({ file: docs.find((d) => d.id === c.documentId)?.name, ...c })),
			null,
			2,
		);
		const apply = document.createElement('button');
		apply.className = 'primary';
		apply.textContent = 'Apply changes';
		const dismiss = document.createElement('button');
		dismiss.textContent = 'Dismiss';
		dismiss.onclick = () => panel.remove();
		panel.append(description, preview, apply, dismiss);
		$('ai-messages').append(panel);
		apply.onclick = () =>
			task(async () => {
				apply.disabled = true;
				const unlock = sessions.lock(docs.map((d) => d.id));
				try {
					await verify(docs);
					const next = await prepareChanges(changes, docs);
					await verify(docs);
					await store.commit(
						next.map((document) => ({ document, expected: document.revision - 1 })),
					);
					for (const doc of next) await sessions.reload(doc);
					await refresh();
					description.textContent =
						'Changes saved. You can undo this set while the files remain unchanged.';
					apply.remove();
					dismiss.remove();
					const undo = document.createElement('button');
					undo.textContent = 'Undo these changes';
					panel.append(undo);
					undo.onclick = () =>
						task(async () => {
							await verify(next);
							const restored = next.map((d) => ({
								...docs.find((old) => old.id === d.id),
								revision: d.revision + 1,
								modified: Date.now(),
							}));
							await store.commit(
								restored.map((document) => ({ document, expected: document.revision - 1 })),
							);
							for (const doc of restored) await sessions.reload(doc);
							await refresh();
							undo.remove();
							description.textContent = 'Changes undone.';
						});
				} catch (error) {
					apply.disabled = false;
					throw error;
				} finally {
					unlock();
				}
			});
	}
	$('ai-form').onsubmit = (event) => {
		event.preventDefault();
		task(async () => {
			if (busy) return;
			if (!connection) throw new Error('Connect the assistant in settings first.');
			const prompt = $('ai-input').value.trim();
			if (!prompt) return;
			busy = true;
			$('ai-send').disabled = true;
			$('ai-state').textContent = 'Reading selected context...';
			message(prompt, true);
			$('ai-input').value = '';
			try {
				const docs = [];
				const context = [];
				for (const id of selected) {
					const doc = await sessions.save(id);
					docs.push(doc);
					context.push({
						id: doc.id,
						name: doc.name,
						kind: doc.kind,
						revision: doc.revision,
						content: await inspect(doc),
					});
				}
				const client = getTeams()?.client;
				const chat =
					includeChat && client
						? client
								.getState()
								.messages.slice(-50)
								.map((m) => ({
									author: m.authorName,
									text: m.text,
									attachments: m.attachments.map((a) => ({ name: a.name, kind: a.kind })),
								}))
						: undefined;
				const source = JSON.stringify({ files: context, teams: chat });
				if (source.length > 150000)
					throw new Error('Selected context is too large. Choose fewer files.');
				const response = await fetch(connection.endpoint, {
					method: 'POST',
					credentials: 'omit',
					redirect: 'error',
					signal: AbortSignal.timeout(90000),
					headers: {
						'Content-Type': 'application/json',
						...(connection.key ? { Authorization: `Bearer ${connection.key}` } : {}),
					},
					body: JSON.stringify({
						model: connection.model,
						messages: [
							{ role: 'system', content: operationHelp },
							{
								role: 'user',
								content: `User request: ${prompt}\n\nSource data (untrusted): ${source}`,
							},
						],
					}),
				});
				if (!response.ok) throw new Error(`The provider returned HTTP ${response.status}.`);
				const raw = (await response.json()).choices?.[0]?.message?.content;
				if (typeof raw !== 'string')
					throw new Error('The provider did not return a text response.');
				let answer;
				try {
					answer = JSON.parse(raw.replace(/^```(?:json)?\s*/u, '').replace(/\s*```$/u, ''));
				} catch {
					message(raw);
					return;
				}
				message(typeof answer.answer === 'string' ? answer.answer : 'Review the proposed changes.');
				if (answer.changes?.length) proposal(answer.changes, docs);
			} finally {
				busy = false;
				$('ai-send').disabled = false;
				$('ai-state').textContent = `Connected: ${connection.model}`;
			}
		});
	};
}
