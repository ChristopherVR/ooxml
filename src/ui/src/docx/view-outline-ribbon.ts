import { group, row, select, stack, tool } from './ribbon-parts';
import type { RibbonIcon } from './ribbon-icons';
import {
	localizeElement,
	translate,
	translateTemplate,
	type EditorLocale,
	type LocalizationKey,
} from './localization';
import type { OutlineAction } from './document-views';
import { refitRibbon } from './ribbon-overflow';
import { panelOf, selectedTab, selectTab, setTabHidden } from './ribbon-tab-api';

export type OutlineKey = Exclude<OutlineAction['key'], 'level'>;

export function mountOutlineRibbon(toolbar: HTMLElement, locale: EditorLocale) {
	const panel = panelOf(toolbar, 'outlining');
	if (!panel) return undefined;
	let optionLocale: EditorLocale | undefined;
	const controls = new Map<
		OutlineKey,
		{ control: HTMLButtonElement | HTMLSelectElement; label: LocalizationKey }
	>();
	const command = (
		key: Exclude<OutlineKey, 'paragraphLevel'>,
		caption: LocalizationKey,
		icon: RibbonIcon,
		label = caption,
		size?: 'inline' | 'large',
	) => {
		const control = tool(
			label,
			icon,
			{ type: 'outline', key },
			{
				caption,
				name: caption,
				inline: size === 'inline',
				large: size === 'large',
			},
		);
		controls.set(key, { control, label });
		return control;
	};
	const levels = Array.from({ length: 9 }, (_, i): [string, string] => [
		String(i + 1),
		String(i + 1),
	]);
	const levelSelect = select('Show Level', [...levels, ['10', 'All Levels']], (value) => ({
		type: 'outline',
		key: 'level',
		value: Number(value),
	}));
	const paragraphLevel = select('Outline Level', [...levels, ['10', 'Body Text']], (value) => ({
		type: 'outline',
		key: 'paragraphLevel',
		value: Number(value),
	}));
	controls.set('paragraphLevel', { control: paragraphLevel, label: 'Outline Level' });
	const show = document.createElement('label');
	show.className = 'dve-outline-filter';
	show.append('Show Level', levelSelect);
	panel.append(
		group(
			'Outline Tools',
			stack(
				row(
					command('promote', 'Promote', 'promote', 'Promote in Outline'),
					paragraphLevel,
					command('demote', 'Demote', 'demote', 'Demote in Outline'),
				),
				row(
					command('up', 'Move Up', 'moveUp', 'Move Up in Outline'),
					command('down', 'Move Down', 'moveDown', 'Move Down in Outline'),
					command('expand', 'Expand', 'next'),
					command('collapse', 'Collapse', 'previous'),
				),
			),
			stack(
				show,
				command('format', 'Show Text Formatting', 'formatPainter', undefined, 'inline'),
				command('firstLine', 'Show First Line Only', 'lineNumbers', undefined, 'inline'),
			),
		),
		group('Close', command('close', 'Close Outline View', 'close', undefined, 'large')),
	);
	localizeElement(panel, locale);
	setTabHidden(toolbar, 'outlining', 'contextual', false);
	selectTab(toolbar, 'outlining');
	refitRibbon(toolbar);

	return {
		sync(
			nextLevel: number,
			selectedLevel: number | undefined,
			formatting: boolean,
			firstLine: boolean,
			enabled: Set<OutlineKey>,
			locale: EditorLocale,
		) {
			levelSelect.value = String(nextLevel);
			paragraphLevel.value = selectedLevel === undefined ? '' : String(selectedLevel);
			if (locale !== optionLocale) {
				optionLocale = locale;
				for (const control of [levelSelect, paragraphLevel]) {
					for (const option of control.options) {
						if (Number(option.value) < 10)
							option.textContent = translateTemplate(locale, 'outline.levelNumber', {
								level: option.value,
							});
					}
				}
			}
			for (const [key, { control, label }] of controls) {
				const pressed = key === 'format' ? formatting : key === 'firstLine' ? firstLine : undefined;
				if (pressed !== undefined) control.setAttribute('aria-pressed', String(pressed));
				control.disabled = !enabled.has(key);
				if (control.disabled) {
					const reason = translate(locale, 'Unavailable for the current selection');
					control.title = reason;
					control.setAttribute('aria-description', reason);
				} else {
					control.removeAttribute('aria-description');
					control.title = translate(locale, label);
				}
			}
		},
		dispose() {
			if (selectedTab(toolbar) === 'outlining') selectTab(toolbar, 'view');
			setTabHidden(toolbar, 'outlining', 'contextual', true);
			panel.replaceChildren();
			refitRibbon(toolbar);
		},
	};
}
