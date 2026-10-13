import type { VisioTextField } from 'ooxml-core/visio';
import { VISIO_FIELD_CATEGORIES } from 'ooxml-core/visio/ui';

export const editableFieldFormats = () => [
	...new Map(
		VISIO_FIELD_CATEGORIES.flatMap((category) => category.formats).map((item) => [
			item.format,
			item,
		]),
	).values(),
];

export function populateExistingFields(
	select: HTMLSelectElement,
	fields: readonly VisioTextField[],
): void {
	const choices = [
		{ value: '', label: 'Insert new field' },
		...fields
			.filter((field) => field.index !== undefined && field.sourceKey !== undefined)
			.map((field) => ({
				value: field.index!,
				label: `Edit field ${Number(field.index) + 1}: ${field.cached.slice(0, 80)}`,
			})),
	];
	select.replaceChildren(
		...choices.map((item) => {
			const option = select.ownerDocument.createElement('option');
			option.value = item.value;
			option.textContent = item.label;
			return option;
		}),
	);
	select.value = '';
}

export function populateExistingFormat(select: HTMLSelectElement, field: VisioTextField): void {
	if ([...select.options].some((option) => option.value === field.format)) {
		select.value = field.format!;
		return;
	}
	const option = select.ownerDocument.createElement('option');
	option.value = '__choose_format__';
	option.textContent = 'Choose a supported format';
	option.disabled = true;
	select.prepend(option);
	select.value = option.value;
}
