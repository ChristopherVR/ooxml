/** The apps, with the maturity each one declares on its own site. */
export const apps = [
	{ id: 'docx', name: 'Word', letter: 'W', color: '#185abd' },
	{ id: 'xlsx', name: 'Excel', letter: 'X', color: '#107c41' },
	{ id: 'pptx', name: 'PowerPoint', letter: 'P', color: '#c43e1c' },
	{ id: 'vsdx', name: 'Visio', letter: 'V', color: '#3955a3', tag: 'Beta' },
];
export const teamsApp = { id: 'teams', name: 'Teams', letter: 'T', color: '#6264a7', tag: 'Beta' };
/** A maturity tag, or nothing for an app without one. */
export const tag = (label) =>
	label ? `<i class="app-tag" data-tone="${label.toLowerCase()}">${label}</i>` : '';
export const $ = (id) => document.getElementById(id);
export const escape = (value) => String(value).replace(/[&<>"']/g, (c) => `&#${c.charCodeAt(0)};`);
export const badge = (kind) => {
	const a = apps.find((a) => a.id === kind);
	return `<span class="app-badge" style="--app:${a?.color ?? teamsApp.color}">${a?.letter ?? teamsApp.letter}</span>`;
};
export function notify(message) {
	$('status').textContent = message instanceof Error ? message.message : message;
}
export function task(fn) {
	return Promise.resolve().then(fn).catch(notify);
}
export function download(doc) {
	const url = URL.createObjectURL(new Blob([doc.bytes]));
	const a = document.createElement('a');
	a.href = url;
	a.download = doc.name;
	a.click();
	setTimeout(() => URL.revokeObjectURL(url), 30000);
}
export function choose(title, content) {
	const dialog = $('dialog');
	$('dialog-title').textContent = title;
	$('dialog-content').replaceChildren(content);
	dialog.showModal();
	return dialog;
}
