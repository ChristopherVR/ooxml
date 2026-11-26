import { bindStencilMenus } from './stencil-menus';

type Menus = ReturnType<typeof bindStencilMenus>;

export function wireStencilMenus(...args: Parameters<typeof bindStencilMenus>) {
	const [pane] = args;
	let loaded: Menus | undefined;
	const menus = () => {
		if (loaded) return loaded;
		loaded = bindStencilMenus(...args);
		pane.append(loaded.menu);
		return loaded;
	};
	const masterButton = (event: Event) =>
		event.target instanceof Element ? event.target.closest<HTMLElement>('[data-master]') : null;
	pane.addEventListener('contextmenu', (event) => {
		if (event.defaultPrevented) return;
		const button = masterButton(event);
		if (button?.isConnected && menus().openMasterMenu(button, event.clientX, event.clientY))
			event.preventDefault();
	});
	pane.addEventListener('keydown', (event) => {
		if (event.defaultPrevented) return;
		if (event.key !== 'ContextMenu' && !(event.shiftKey && event.key === 'F10')) return;
		const button = masterButton(event);
		if (!button?.isConnected) return;
		const box = button.getBoundingClientRect();
		if (menus().openMasterMenu(button, box.left, box.bottom)) event.preventDefault();
	});
	return {
		create: () => menus().create(),
		openSection: (button: HTMLElement, id: string) => menus().openSection(button, id),
		close: (id: string) => loaded?.close(id),
	};
}
