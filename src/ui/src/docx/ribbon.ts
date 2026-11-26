import { defineRibbon } from '../controls';
import { shortcutHint } from './binding-labels';
import { emit } from './events';
import { setComboValue, type ComboInput } from './ribbon-combo';
import {
	localeOf,
	loadTableLocalization,
	localizeElement,
	normalizeEditorLocale,
	translate,
	type LocalizationKey,
} from './localization';
import type { RibbonAction } from './ribbon-action';
import { buildHomePanel } from './ribbon-home';
import { buildNativeTabs } from './ribbon-native-tabs';
import { buildMailings } from './mailings-ribbon';
import { attachRibbonBehavior } from './ribbon-behavior';
import { attachRibbonOverflow, refitRibbon } from './ribbon-overflow';
import { buildOtherPanels } from './ribbon-tabs';
import { buildHeaderFooterPanel } from './header-footer-ribbon';
import { buildArrangeGroup, textDirectionMenu } from './arrange-ribbon';
import { group } from './ribbon-parts';
import { buildShapeFormatPanel } from './shape-format-ribbon';
import { buildPictureFormatPanel } from './picture-format-ribbon';
import { buildSmartArtDesignPanel } from './smartart-design-ribbon';
import { syncSmartArtFormatGalleries } from './smartart-format-galleries';
import { buildChartDesignPanel } from './chart-design-ribbon';
import { panelOf, panelsOf, setTabHidden } from './ribbon-tab-api';
import { preloadDesignGallery } from './design-ribbon-controls';
import { TABLE_TABS } from './table-ribbon-keys';
import { reapplyRibbonVisibility } from './ribbon-visibility';

export type { RibbonAction } from './ribbon-action';

export function createRibbon(locale: string = 'en'): HTMLElement {
	defineRibbon();
	const root = document.createElement('office-ui-ribbon');
	root.className = 'dve-ribbon';
	root.setAttribute('role', 'toolbar');
	root.setAttribute('aria-label', 'Document formatting');
	root.setAttribute('label', 'Ribbon tabs');
	root.setAttribute('file-label', 'File');
	root.setAttribute('file-expanded', 'false');
	root.setAttribute('collapse-label', 'Collapse the ribbon');
	const panels = new Map<string, HTMLElement>();
	for (const name of [
		'Home',
		'Insert',
		'Design',
		'Layout',
		'References',
		'Mailings',
		'Review',
		'View',
		'Help',
		'Table Design',
		'Table Layout',
		'Header & Footer',
		'Shape Format',
		'Picture Format',
		'SmartArt Design',
		'Chart Design',
		'Outlining',
	]) {
		const key =
			name === 'Header & Footer' ? 'header-footer' : name.toLowerCase().replace(/ /g, '-');
		const panel = document.createElement('div');
		panel.className = 'ribbon-panel';
		panel.dataset.ribbonTab = key;
		panel.dataset.label = name;
		panel.dataset.panel = name;
		panel.id = `dve-panel-${key}`;
		panel.tabIndex = 0;
		if (
			[
				'Table Design',
				'Table Layout',
				'Header & Footer',
				'Shape Format',
				'Picture Format',
				'Chart Design',
				'Outlining',
				'SmartArt Design',
			].includes(name)
		)
			panel.dataset.contextual = '';
		panels.set(name, panel);
	}
	buildHomePanel(panels);
	buildOtherPanels(panels);
	buildNativeTabs(panels);
	panels.get('Layout')!.append(buildArrangeGroup(), group('Text', textDirectionMenu()));
	// Customize Ribbon and Tell me discover the controls before their workflow is loaded.
	buildMailings(panels.get('Mailings')!);
	buildHeaderFooterPanel(panels.get('Header & Footer')!);
	buildShapeFormatPanel(panels.get('Shape Format')!);
	buildPictureFormatPanel(panels.get('Picture Format')!);
	buildSmartArtDesignPanel(panels.get('SmartArt Design')!);
	buildChartDesignPanel(panels.get('Chart Design')!);
	for (const panel of panels.values()) root.append(panel);
	// Table tools are contextual: the tab appears only while the selection is in a table.
	for (const key of [
		'outlining',
		...TABLE_TABS,
		'header-footer',
		'shape-format',
		'picture-format',
		'chart-design',
		'smartart-design',
	])
		setTabHidden(root, key, 'contextual', true);
	setComboValue(root.querySelector<ComboInput>('[aria-label="Font family"]')!, 'Calibri');
	setComboValue(root.querySelector<ComboInput>('[aria-label="Font size"]')!, '11');
	root.querySelector<HTMLSelectElement>('[aria-label="Zoom"]')!.value = '100';
	root.querySelector<HTMLSelectElement>('[aria-label="Layout view"]')!.value = 'draft';
	attachRibbonBehavior(root);
	attachRibbonOverflow(root);
	root.addEventListener('office-ribbon-select', (event) => {
		if (event instanceof CustomEvent && event.detail?.tab === 'design') preloadDesignGallery();
	});
	root.addEventListener('click', (event) => {
		const target = (event.target as HTMLElement).closest<HTMLButtonElement>('button[data-action]');
		if (!target) return;
		emit(target, 'ribbon-action', JSON.parse(target.dataset.action!) as RibbonAction);
	});
	for (const type of ['pointerover', 'focusin', 'change'])
		root.addEventListener(type, (event) => {
			if (event.target instanceof HTMLSelectElement) syncSelectTitles(event.target.parentElement!);
		});
	setRibbonLocale(root, locale);
	void mountInsertPartsRibbon(root);
	return root;
}

/** Gives every select the tooltip of its selected label so an ellipsized value stays readable. */
export function syncSelectTitles(root: ParentNode): void {
	for (const select of root.querySelectorAll('select')) {
		const text = select.selectedOptions[0]?.textContent?.trim() ?? '';
		if (text && select.title !== text) select.title = text;
	}
}

/** Gives every text button its full label as a tooltip, for when a narrow slot truncates it. */
function syncButtonTitles(root: ParentNode): void {
	for (const button of root.querySelectorAll('button')) {
		if (button.hasAttribute('data-unavailable')) continue;
		const id = button.dataset.localearialabel;
		// A command named apart from its label keeps the label as its tooltip.
		const label =
			button.dataset.localename && id
				? translate(localeOf(button), id as LocalizationKey)
				: button.getAttribute('aria-label');
		if (!label) continue;
		const hint = shortcutHint(id ?? label);
		const title = hint ? `${label} (${hint})` : label;
		if ((button.textContent?.trim() || hint) && button.title !== title) button.title = title;
	}
}

export function mountInsertPartsRibbon(root: HTMLElement): Promise<void> {
	return import('./insert-parts-ribbon').then((module) => module.mountInsertParts(root));
}

export function setRibbonLocale(root: HTMLElement, value: string): void {
	const locale = normalizeEditorLocale(value);
	root.dataset.editorLocale = locale;
	localizeElement(root, locale);
	syncSmartArtFormatGalleries(root);
	// The tab row is drawn by the shared ribbon from the panels' labels, so translate those.
	// A host's add-in tab keeps the label the host gave it.
	for (const panel of panelsOf(root))
		if (!('addIn' in panel.dataset))
			panel.dataset.label = translate(locale, `tab.${panel.dataset.ribbonTab}` as LocalizationKey);
	root.setAttribute('file-label', translate(locale, 'File'));
	root.setAttribute('collapse-label', translate(locale, 'Collapse the ribbon'));
	syncSelectTitles(root);
	syncButtonTitles(root);
	refitRibbon(root);
}

/** Load contextual commands after the selection first enters a table. */
export async function loadTablePanels(root: HTMLElement) {
	const [{ buildTablePanels, syncTableControls }] = await Promise.all([
		import('./table-ribbon'),
		loadTableLocalization(),
	]);
	const design = panelOf(root, 'table-design');
	const layout = panelOf(root, 'table-layout');
	if (design && layout) {
		buildTablePanels(design, layout);
		setRibbonLocale(root, localeOf(root));
		reapplyRibbonVisibility(root);
	}
	return syncTableControls;
}
