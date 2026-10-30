import type { RibbonAction } from './ribbon-action';
import type { UnderlineKind } from './font-format';
import { splitInline, tool } from './ribbon-parts';
import type { RibbonIcon } from './ribbon-icons';

/** Split commands retain the default action while exposing Word's related choices. */
function split(
	main: HTMLButtonElement,
	label: string,
	choices: Array<[string, string]>,
	action: (value: string) => RibbonAction,
	large = false,
): HTMLElement {
	const control = splitInline(main, label, choices, action);
	control.classList.add('ribbon-command-split');
	if (large) control.classList.add('ribbon-command-split-large');
	return control;
}

export function pasteSplit(): HTMLElement {
	return split(
		tool('Paste', 'paste', { type: 'clipboard', key: 'paste' }, { large: true }),
		'Paste options',
		[
			['paste', 'Paste'],
			['text', 'Keep text only'],
		],
		(value) => ({ type: 'clipboard', key: 'paste', plainText: value === 'text' }),
		true,
	);
}

export function underlineSplit(): HTMLElement {
	return split(
		tool('Underline', 'underline', { type: 'format', key: 'underline' }),
		'Underline style',
		[
			['single', 'Single'],
			['double', 'Double'],
			['dotted', 'Dotted'],
			['dash', 'Dashed'],
			['wave', 'Wave'],
			['none', '(none)'],
		],
		(value) => ({ type: 'underline', value: value as UnderlineKind }),
	);
}

export function listSplit(key: 'bullet' | 'number'): HTMLElement {
	const label = key === 'bullet' ? 'Bulleted list' : 'Numbered list';
	const icon: RibbonIcon = key === 'bullet' ? 'bullets' : 'numbering';
	return split(
		tool(label, icon, { type: 'list', key }),
		key === 'bullet' ? 'Bullets options' : 'Numbering options',
		[
			[key, label],
			...(key === 'bullet' ? [['picture', 'Picture bullet...'] as [string, string]] : []),
			['remove', 'Remove list'],
		],
		(value) =>
			value === 'picture'
				? { type: 'pictureBullet' }
				: { type: 'list', key: value as 'bullet' | 'number' | 'remove' },
	);
}
