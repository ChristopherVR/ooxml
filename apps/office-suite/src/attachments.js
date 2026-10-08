import { workspaceUrl, workspaceId } from './product.js';
import { SuiteAttachmentStore } from 'ooxml-ui/suite';
import { documentDatabase } from './profile-state.js';
import { choose, download, escape } from './ui.js';

export const attachments = new SuiteAttachmentStore(`${documentDatabase()}-attachments`);
export const attachmentUrl = (id) => workspaceUrl('attachment', id);
export const localAttachmentId = (raw) => workspaceId(raw, 'attachment');
export async function showAttachment(id) {
	const file = await attachments.get(id);
	const content = document.createElement('div');
	content.className = 'attachment-preview';
	content.innerHTML = `<p>${escape(file.name)} · ${(file.bytes.length / 1024).toFixed(1)} KB</p>`;
	// Only passive media gets an inline preview. HTML, SVG and executable content download as files.
	const tag = /^image\/(png|jpeg|gif|webp)$/u.test(file.mime)
		? 'img'
		: /^audio\/(mpeg|ogg|wav|mp4|webm)$/u.test(file.mime)
			? 'audio'
			: /^video\/(mp4|webm|ogg)$/u.test(file.mime)
				? 'video'
				: null;
	let url;
	if (tag) {
		url = URL.createObjectURL(new Blob([file.bytes], { type: file.mime }));
		const media = document.createElement(tag);
		media.src = url;
		if (tag === 'img') media.alt = file.name;
		else media.controls = true;
		content.append(media);
	} else
		content.insertAdjacentHTML(
			'beforeend',
			'<p>Download this file to open it in its application.</p>',
		);
	const button = document.createElement('button');
	button.className = 'primary';
	button.textContent = 'Download file';
	button.onclick = () => download(file);
	content.append(button);
	const dialog = choose('Attachment', content);
	dialog.addEventListener(
		'close',
		() => {
			if (url) URL.revokeObjectURL(url);
		},
		{ once: true },
	);
}
