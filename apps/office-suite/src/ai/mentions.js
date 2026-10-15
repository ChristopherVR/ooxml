/**
 * `@` mentions in the composer: typing `@` lists the open tabs first, then the library, filtered
 * as the user types. Picking one writes `@name` into the message and attaches the file.
 */
import { badge, escape } from '../ui.js';

export function mountMentions({ input, items, onPick }) {
	const list = document.createElement('div');
	list.className = 'chat-mentions';
	list.setAttribute('role', 'listbox');
	list.setAttribute('aria-label', 'Files to mention');
	list.hidden = true;
	input.closest('form').prepend(list);
	let matches = [],
		index = 0,
		start = -1;

	const close = () => {
		list.hidden = true;
		start = -1;
		input.removeAttribute('aria-activedescendant');
	};
	function draw() {
		list.innerHTML = matches.length
			? matches
					.map(
						(m, i) =>
							`${i === 0 || m.group !== matches[i - 1].group ? `<span class="nav-caption">${m.group}</span>` : ''}<button type="button" role="option" id="mention-${i}" aria-selected="${i === index}" data-mention="${i}">${badge(m.kind)}<span>${escape(m.name)}</span></button>`,
					)
					.join('')
			: '<p>No matching files</p>';
		input.setAttribute('aria-activedescendant', `mention-${index}`);
		list.querySelector('[aria-selected=true]')?.scrollIntoView({ block: 'nearest' });
	}
	async function update() {
		const caret = input.selectionStart;
		const before = input.value.slice(0, caret);
		const at = before.search(/(^|\s)@[^\s@]*$/u);
		if (at < 0) return close();
		start = before.indexOf('@', at);
		const query = before.slice(start + 1).toLowerCase();
		matches = (await items()).filter((m) => m.name.toLowerCase().includes(query)).slice(0, 12);
		index = 0;
		list.hidden = false;
		draw();
	}
	function pick(i) {
		const match = matches[i];
		if (!match || start < 0) return;
		const caret = input.selectionStart;
		input.value = `${input.value.slice(0, start)}@${match.name} ${input.value.slice(caret)}`;
		const next = start + match.name.length + 2;
		input.setSelectionRange(next, next);
		close();
		onPick(match);
		input.dispatchEvent(new Event('input'));
		input.focus();
	}
	input.addEventListener('input', () => void update());
	input.addEventListener('blur', () => setTimeout(close, 150));
	input.addEventListener(
		'keydown',
		(event) => {
			if (list.hidden || !matches.length) {
				if (event.key === 'Escape') close();
				return;
			}
			if (event.key === 'ArrowDown' || event.key === 'ArrowUp') {
				event.preventDefault();
				index = (index + (event.key === 'ArrowDown' ? 1 : -1) + matches.length) % matches.length;
				draw();
			} else if (event.key === 'Enter' || event.key === 'Tab') {
				event.preventDefault();
				event.stopImmediatePropagation();
				pick(index);
			} else if (event.key === 'Escape') {
				event.preventDefault();
				close();
			}
		},
		{ capture: true },
	);
	list.addEventListener('mousedown', (event) => event.preventDefault());
	list.addEventListener('click', (event) => {
		const i = event.target.closest('[data-mention]')?.dataset.mention;
		if (i !== undefined) pick(Number(i));
	});
}
