/**
 * Reordering the document tabs: drag a tab along the strip, or press Ctrl+Shift+Left/Right on it.
 * A dragged tab also carries its document id, so it can be dropped on a chat to attach it.
 */
import { TAB_TYPE } from './ai/attach.js';

export function mountTabOrder({ strip, sessions }) {
	let dragging = null;
	const fileTabs = () => [...strip.querySelectorAll('.file-tab')];
	const clear = () => {
		for (const tab of fileTabs()) tab.classList.remove('drop-before', 'drop-after');
	};
	/** Where a tab dropped at `x` lands: its index among the file tabs. */
	const slot = (x) => {
		const tabs = fileTabs().filter((tab) => tab.dataset.tab !== dragging);
		const at = tabs.findIndex((tab) => {
			const box = tab.getBoundingClientRect();
			return x < box.left + box.width / 2;
		});
		return { tabs, index: at < 0 ? tabs.length : at };
	};
	strip.addEventListener('dragstart', (event) => {
		const tab = event.target.closest?.('.file-tab');
		if (!tab) return;
		dragging = tab.dataset.tab;
		event.dataTransfer.setData(TAB_TYPE, dragging);
		event.dataTransfer.setData('text/plain', tab.textContent.replace(/×$/u, '').trim());
		event.dataTransfer.effectAllowed = 'copyMove';
		tab.classList.add('dragging');
	});
	strip.addEventListener('dragend', () => {
		strip.querySelector('.dragging')?.classList.remove('dragging');
		dragging = null;
		clear();
	});
	strip.addEventListener('dragover', (event) => {
		if (!dragging) return;
		event.preventDefault();
		event.dataTransfer.dropEffect = 'move';
		clear();
		const { tabs, index } = slot(event.clientX);
		if (tabs[index]) tabs[index].classList.add('drop-before');
		else tabs.at(-1)?.classList.add('drop-after');
	});
	strip.addEventListener('dragleave', (event) => {
		if (!strip.contains(event.relatedTarget)) clear();
	});
	strip.addEventListener('drop', (event) => {
		if (!dragging) return;
		event.preventDefault();
		event.stopPropagation();
		const { index } = slot(event.clientX);
		const id = dragging;
		clear();
		sessions.move(id, index);
		strip.querySelector(`[data-tab="${CSS.escape(id)}"] [role=tab]`)?.focus();
	});
	strip.addEventListener('keydown', (event) => {
		if (!event.ctrlKey || !event.shiftKey || !['ArrowLeft', 'ArrowRight'].includes(event.key))
			return;
		const id = event.target.closest('.file-tab')?.dataset.tab;
		if (!id) return;
		event.preventDefault();
		event.stopImmediatePropagation();
		const order = [...sessions.sessions.keys()];
		sessions.move(id, order.indexOf(id) + (event.key === 'ArrowRight' ? 1 : -1));
		strip.querySelector(`[data-tab="${CSS.escape(id)}"] [role=tab]`)?.focus();
	});
}
