import { connectOfficeAddIn } from '../../src/ui/src/ribbon/add-in-client';

const parentOrigin = document.currentScript?.getAttribute('data-editor-origin');
if (!parentOrigin) throw new Error('Missing fixture editor origin.');
const Office = connectOfficeAddIn(parentOrigin);
const ready = document.querySelector('#ready');
const status = document.querySelector('#result');
const isolated = document.querySelector('#isolated');
try {
	if (isolated) isolated.textContent = window.parent.document.title;
} catch {
	if (isolated) isolated.textContent = 'Editor DOM is isolated';
}
void Office.onReady(({ host }) => {
	if (ready) ready.textContent = `${host} / ${Office.context.host}`;
});

document.querySelector('#read')?.addEventListener('click', () => {
	void Office.context.document.getSelectedDataAsync('text', null, (result) => {
		if (status)
			status.textContent =
				result.status === 'succeeded' ? `Read: ${result.value || '(empty)'}` : result.error.message;
	});
});

document.querySelector('#write')?.addEventListener('click', () => {
	void Office.context.document.setSelectedDataAsync('Updated by add-in', null, (result) => {
		if (status)
			status.textContent = result.status === 'succeeded' ? 'Wrote text' : result.error.message;
	});
});
