// The equation editor's messages: why an equation stays as it is, and the box's name, which says
// which equation of a display paragraph with several it edits.
import type { EquationReadOnlyReason } from 'ooxml-core/math';
import {
	translate,
	translateTemplate,
	type EditorLocale,
	type LocalizationKey,
} from './localization';

export const NO_LINEAR: LocalizationKey =
	'This equation cannot be shown in Linear without changing it, so it is kept as it is.';
export const READ_ONLY: Record<EquationReadOnlyReason, LocalizationKey> = {
	tracked:
		'This equation has tracked changes its linear text cannot hold, so it is kept as it is. Accept or reject them to edit it.',
	formatting: 'This equation mixes formatting its linear text cannot hold, so it is kept as it is.',
	content: 'This equation has content its linear text cannot hold, so it is kept as it is.',
};

/** The editor's accessible name and, for a paragraph with several equations, "Equation 2 of 3". */
export function editorLabel(
	locale: EditorLocale,
	part: number,
	parts: number,
): { name: string; part: string } {
	const name = translate(locale, 'Equation editor');
	if (parts < 2) return { name, part: '' };
	const counter = translateTemplate(locale, 'Equation {n} of {count}', {
		n: part + 1,
		count: parts,
	});
	return { name: `${name}, ${counter}`, part: counter };
}
