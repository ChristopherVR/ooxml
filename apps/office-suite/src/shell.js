import { $ } from './ui.js';
import { icon } from './icons.js';
import { shellLayout } from './shell-layout.js';

/**
 * Adapt the shell to phones: size it to the visible viewport (address bar, on-screen keyboard),
 * fold the chrome away in landscape and offer the document actions as one More menu.
 * `state()` returns `{ view, documents, teamsOpen }` for the current render; the returned
 * function re-applies the layout after the state changes.
 */
export function mountShell(state) {
	let expanded = false;
	const toggle = $('chrome-toggle');
	toggle.innerHTML = icon('chevron');
	$('actions-more').innerHTML = icon('more');
	const measure = () => {
		const viewport = window.visualViewport;
		const height = Math.round(viewport?.height ?? window.innerHeight);
		document.documentElement.style.setProperty('--app-h', `${height}px`);
		return { width: Math.round(viewport?.width ?? window.innerWidth), height };
	};
	const apply = () => {
		const layout = shellLayout({ ...measure(), ...state(), expanded });
		const body = document.body;
		body.dataset.narrow = String(layout.narrow);
		body.dataset.hideTabs = String(layout.hideTabs);
		body.dataset.collapsible = String(layout.collapsible);
		body.dataset.chrome = layout.collapsed ? 'collapsed' : 'full';
		toggle.setAttribute('aria-expanded', String(!layout.collapsed));
	};
	toggle.addEventListener('click', () => {
		expanded = !expanded;
		apply();
	});
	window.addEventListener('resize', apply);
	window.visualViewport?.addEventListener('resize', apply);
	mountMoreMenu();
	apply();
	return apply;
}

function mountMoreMenu() {
	const menu = $('actions-menu');
	const more = $('actions-more');
	more.addEventListener('click', () => {
		const items = [...document.querySelectorAll('.document-commandbar [data-menu]')]
			.filter((button) => !button.hidden)
			.map((button) => {
				const item = document.createElement('button');
				item.type = 'button';
				item.setAttribute('role', 'menuitem');
				item.textContent = button.textContent;
				item.addEventListener('click', () => {
					menu.hidePopover();
					button.click();
				});
				return item;
			});
		menu.replaceChildren(...items);
		const rect = more.getBoundingClientRect();
		menu.style.inset = `${Math.round(rect.bottom + 4)}px ${Math.max(8, Math.round(window.innerWidth - rect.right))}px auto auto`;
		menu.showPopover();
		items[0]?.focus();
	});
}
