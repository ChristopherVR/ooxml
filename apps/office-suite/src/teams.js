import { workspaceUrl, workspaceId } from './product.js';
import { notify, download } from './ui.js';
import { uploadTeamsServerFile, readContent } from 'ooxml-core/teams';
import { onTheme, currentTheme, setThemeMode } from './themes.js';
import { attachments, attachmentUrl, localAttachmentId, showAttachment } from './attachments.js';
import { applyOfficeTheme } from './theme-tokens.js';
import { currentProfile, teamsWorkspace, onProfile } from './profile-state.js';

export const documentUrl = (id) => workspaceUrl('file', id);
export const localDocumentId = (raw) => workspaceId(raw, 'file');
export async function mountTeams(host, store, open, refresh) {
	const { defineTeamsApp } = await import('ooxml-ui/teams');
	defineTeamsApp();
	const teams = document.createElement('teams-app');
	const profile = currentProfile();
	teams.toggleAttribute('host-profile', true);
	teams.setAttribute('host-theme', '');
	teams.addEventListener('teams-settings-theme', (event) => setThemeMode(event.detail.theme));
	teams.workspaceId = teamsWorkspace();
	teams.userName = profile.name;
	teams.userId = profile.userId;
	teams.uploadFile = async (file, context) => {
		if (teams.client?.getState().mode === 'server')
			return uploadTeamsServerFile(
				teams.client.workspace.config,
				context.workspaceId,
				file,
				context,
			);
		if (!(file instanceof Blob)) throw new Error('File bytes are required.');
		if (file.size > 32 * 1024 * 1024) throw new Error('Choose an attachment smaller than 32 MB.');
		const bytes = new Uint8Array(await file.arrayBuffer());
		const name = file.originalName || file.name;
		if (!/\.(docx|xlsx|pptx|vsdx)$/iu.test(name)) {
			const item = await attachments.create(name, file.type || 'application/octet-stream', bytes);
			return { url: attachmentUrl(item.id) };
		}
		const doc = await store.create(name, bytes);
		await refresh();
		return { url: documentUrl(doc.id) };
	};
	teams.addEventListener('teams-file-transfer', (event) => {
		const raw = event.detail.attachment.url;
		const generic = localAttachmentId(raw);
		const id = localDocumentId(raw);
		if (!id && !generic) return;
		event.preventDefault();
		void (async () => {
			if (event.detail.action === 'download')
				download(generic ? await attachments.get(generic) : await store.get(id));
			else {
				await navigator.clipboard.writeText(generic ? attachmentUrl(generic) : documentUrl(id));
				notify('Local workspace link copied');
			}
		})().catch(notify);
	});
	teams.addEventListener('teams-open-file', (event) => {
		event.preventDefault();
		const { url, attachment, channelId } = event.detail;
		void (async () => {
			const sourceUrl = attachment.url ?? url;
			const generic = localAttachmentId(sourceUrl);
			if (generic) {
				await showAttachment(generic);
				return;
			}
			let id = localDocumentId(sourceUrl);
			if (!id) {
				const existing = (await store.list()).find((d) => d.sourceUrl === sourceUrl);
				if (existing) id = existing.id;
				else {
					if (!url || !/^https?:$/u.test(new URL(url).protocol))
						throw new Error('Attachment URL is unavailable.');
					const bytes = await readContent(url, AbortSignal.timeout(60000), 32 * 1024 * 1024);
					if (!/\.(docx|xlsx|pptx|vsdx)$/iu.test(attachment.name)) {
						const file = await attachments.create(
							attachment.name,
							attachment.mime || 'application/octet-stream',
							bytes,
						);
						await showAttachment(file.id);
						return;
					}
					const doc = await store.create(attachment.name, bytes, { sourceUrl });
					id = doc.id;
				}
			}
			await open(id, { channelId, remote: !localDocumentId(url) });
		})().catch(notify);
	});
	host.append(teams);
	await teams.updateComplete;
	if (teams.client?.getState().mode === 'local' && !teams.client.workspace.chat.channels().length)
		teams.client.workspace.chat.createChannel({ id: 'general', name: 'General' });
	await new Promise((resolve) => queueMicrotask(resolve));
	onTheme(({ mode, accent }) => {
		teams.theme = mode;
		applyOfficeTheme(teams, { mode, accent });
	});
	onProfile((p) => {
		if (teams.userName === p.name) return;
		teams.client?.flushStorage();
		teams.userName = p.name;
		void teams.updateComplete.then(() => {
			teams.theme = currentTheme().mode;
		});
	});
	return teams;
}

/** A local reference keeps pointing to the same document after edits, never to an imported copy. */
export async function shareDocument(teams, doc, channelId) {
	const client = teams.client;
	if (!client) throw new Error('Open Teams before sharing.');
	if (client.getState().mode !== 'local') {
		await client.saveFileCopy(channelId, new File([doc.bytes], doc.name));
		return;
	}
	const posted = client.workspace.chat.post(channelId, {
		text: `Shared ${doc.name}`,
		attachments: [
			{ name: doc.name, kind: doc.kind, size: doc.bytes.length, url: documentUrl(doc.id) },
		],
	});
	if (!posted) throw new Error('The file could not be shared.');
	client.flushStorage();
}
