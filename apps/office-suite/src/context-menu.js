/** Small keyboard-accessible context menu, shared by tabs and library rows. */
export function contextMenu(event, items) {
	event.preventDefault();
	document.querySelector('.workspace-context-menu')?.remove();
	const menu = document.createElement('div');
	menu.className = 'workspace-context-menu';
	menu.setAttribute('popover', 'auto');
	menu.setAttribute('role', 'menu');
	for (const item of items) {
		const button = document.createElement('button');
		button.type = 'button';
		button.setAttribute('role', 'menuitem');
		button.textContent = item.label;
		button.disabled = !!item.disabled;
		button.onclick = () => {
			menu.hidePopover();
			Promise.resolve()
				.then(item.run)
				.catch(item.error ?? console.error);
		};
		menu.append(button);
	}
	const source =
		event.currentTarget instanceof HTMLElement ? event.currentTarget : document.activeElement;
	menu.addEventListener('toggle', (e) => {
		if (e.newState === 'closed') {
			menu.remove();
			if (source?.isConnected) source.focus({ preventScroll: true });
		}
	});
	menu.addEventListener('keydown', (e) => {
		const buttons = [...menu.querySelectorAll('button:not(:disabled)')];
		const index = buttons.indexOf(document.activeElement);
		if (['ArrowDown', 'ArrowUp', 'Home', 'End'].includes(e.key)) {
			e.preventDefault();
			buttons[
				e.key === 'Home'
					? 0
					: e.key === 'End'
						? buttons.length - 1
						: (index + (e.key === 'ArrowDown' ? 1 : -1) + buttons.length) % buttons.length
			]?.focus();
		}
	});
	document.body.append(menu);
	menu.showPopover();
	const anchor = event.target.getBoundingClientRect();
	const x = event.clientX || anchor.left,
		y = event.clientY || anchor.bottom;
	menu.style.left = `${Math.max(8, Math.min(x, innerWidth - menu.offsetWidth - 8))}px`;
	menu.style.top = `${Math.max(8, Math.min(y, innerHeight - menu.offsetHeight - 8))}px`;
	menu.querySelector('button:not(:disabled)')?.focus();
}
