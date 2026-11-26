import { loadThemePresets } from 'ooxml-core/drawingml';
import type { EditorCore } from './editor-core';
import { translate } from './localization';
import { tool } from './ribbon-parts';
import { selectedTab } from './ribbon-tab-api';
import type { RibbonIcon } from './ribbon-icons';

export type DesignMenu = 'theme' | 'colors' | 'fonts' | 'spacing' | 'effects' | 'default';

export function designControl(
	label: string,
	icon: RibbonIcon,
	menu: DesignMenu,
): HTMLButtonElement {
	const button = tool(label, icon, { type: 'documentDesign', menu }, { large: true });
	button.dataset.design = menu;
	if (menu !== 'default') {
		button.setAttribute('aria-haspopup', 'menu');
		button.setAttribute('aria-expanded', 'false');
		button.addEventListener('keydown', (event) => {
			if (event.key !== 'ArrowDown' || button.getAttribute('aria-expanded') === 'true') return;
			event.preventDefault();
			button.click();
		});
	}
	return button;
}

export function canEditDesign(core: EditorCore): boolean {
	return Boolean(
		!core.loading &&
		core.view?.editable &&
		core.host.canEditOutsideBody() &&
		!core.model.protection?.enforced,
	);
}

export function syncDesignControls(core: EditorCore, toolbar: HTMLElement): void {
	for (const button of toolbar.querySelectorAll<HTMLButtonElement>('button[data-design]')) {
		const noStore = button.dataset.design === 'default' && !core.wordDesignPreferences;
		button.disabled = noStore || !canEditDesign(core);
		const reason = button.disabled
			? translate(core.locale, noStore ? 'design.noPreferences' : 'design.unavailable')
			: undefined;
		if (reason) button.setAttribute('aria-description', reason);
		else button.removeAttribute('aria-description');
		button.title = reason ?? button.getAttribute('aria-label') ?? '';
	}
}

const loadDesignGallery = () => Promise.all([import('./design-gallery'), loadThemePresets()]);

export async function openDesignMenu(core: EditorCore, menu: DesignMenu): Promise<void> {
	const toolbar = core.shell.toolbar;
	const anchor = toolbar?.querySelector<HTMLButtonElement>(`button[data-design="${menu}"]`);
	if (!toolbar || !anchor || anchor.disabled) return;
	const [{ openDocumentDesign }, presets] = await loadDesignGallery();
	if (anchor.isConnected && !anchor.disabled && selectedTab(toolbar) === 'design')
		openDocumentDesign(core, anchor, menu, presets);
}
