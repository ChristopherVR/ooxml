import { loadStencilStrings, translateStencilText } from './stencil-locale';
import { stencilErrorMessage } from './stencil-files';
import type { bindStencilMenus } from './stencil-menus';

type Menus = ReturnType<typeof bindStencilMenus>;

export function lazyStencilMenus(...args: Parameters<typeof bindStencilMenus>) {
	const [pane, , , options] = args;
	let pending: Promise<Menus> | undefined;
	let loaded: Menus | undefined;
	const use = (action: (menus: Menus) => void) => {
		if (loaded) {
			action(loaded);
			return;
		}
		pending ??= Promise.all([import('./stencil-menus'), loadStencilStrings()])
			.then(([{ bindStencilMenus }]) => {
				loaded = bindStencilMenus(...args);
				pane.append(loaded.menu);
				return loaded;
			})
			.catch((error: unknown) => {
				pending = undefined;
				throw error;
			});
		void pending.then(action, (error: unknown) =>
			options.error(stencilErrorMessage(error, (text) => translateStencilText(pane, text))),
		);
	};
	const masterButton = (event: Event) =>
		event.target instanceof Element ? event.target.closest<HTMLElement>('[data-master]') : null;
	const openMaster = (button: HTMLElement, x: number, y: number) => {
		use((menus) => {
			if (button.isConnected) menus.openMasterMenu(button, x, y);
		});
	};
	pane.addEventListener('contextmenu', (event) => {
		const button = masterButton(event);
		if (!button) return;
		event.preventDefault();
		openMaster(button, event.clientX, event.clientY);
	});
	pane.addEventListener('keydown', (event) => {
		if (event.key !== 'ContextMenu' && !(event.shiftKey && event.key === 'F10')) return;
		const button = masterButton(event);
		if (!button) return;
		event.preventDefault();
		const box = button.getBoundingClientRect();
		openMaster(button, box.left, box.bottom);
	});
	return { use, close: (id: string) => loaded?.close(id) };
}
