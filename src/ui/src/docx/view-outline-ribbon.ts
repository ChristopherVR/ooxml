import { button, group, menuSelect, stack } from './ribbon-parts';
import {
	localizeElement,
	translate,
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
	const controls = new Map<OutlineKey, { control: HTMLButtonElement; label: LocalizationKey }>();
	const command = (key: OutlineKey, caption: LocalizationKey, label = caption) => {
		const control = button(label, caption, { type: 'outline', key });
		control.dataset.localearialabel = label;
		control.dataset.localename = caption;
		controls.set(key, { control, label });
		return control;
	};
	const level = menuSelect(
		'Show Level',
		'view',
		[
			...Array.from({ length: 9 }, (_, i): [string, string] => [String(i + 1), String(i + 1)]),
			['10', 'All Levels'],
		],
		(value) => ({ type: 'outline', key: 'level', value: Number(value) }),
	);
	const levelSelect = level.querySelector('select');
	if (levelSelect) levelSelect.value = '10';
	panel.append(
		group(
			'Outline',
			level,
			stack(
				command('promote', 'Promote', 'Promote in Outline'),
				command('demote', 'Demote', 'Demote in Outline'),
			),
			stack(
				command('up', 'Move Up', 'Move Up in Outline'),
				command('down', 'Move Down', 'Move Down in Outline'),
			),
			stack(command('expand', 'Expand'), command('collapse', 'Collapse')),
			stack(
				command('format', 'Show Text Formatting'),
				command('firstLine', 'Show First Line Only'),
			),
			command('close', 'Close Outline View'),
		),
	);
	localizeElement(panel, locale);
	setTabHidden(toolbar, 'outlining', 'contextual', false);
	selectTab(toolbar, 'outlining');
	refitRibbon(toolbar);

	return {
		sync(
			nextLevel: number,
			formatting: boolean,
			firstLine: boolean,
			enabled: Set<OutlineKey>,
			locale: EditorLocale,
		) {
			if (levelSelect) levelSelect.value = String(nextLevel);
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
