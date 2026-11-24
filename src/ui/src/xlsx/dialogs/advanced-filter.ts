// Data > Advanced Filter: filter a list in place or copy matching rows using a criteria range.
import {
	formatRange,
	type AdvancedFilterAction,
	type AdvancedFilterOptions,
	type AdvancedFilterResult,
} from 'ooxml-core/xlsx';
import { regionOf, target } from 'ooxml-core/xlsx/ui';
import type { EditorContext } from 'ooxml-core/xlsx/ui';
import { checkbox, field, invalid, radios, textInput } from './fields';
import { showDialog } from './frame';
import { resolveReference } from './go-to';

export function openAdvancedFilter(ctx: EditorContext): Promise<AdvancedFilterResult | undefined> {
	const t = target(ctx);
	if (!t) return Promise.resolve(undefined);
	const region = regionOf(t);

	const action = radios(
		ctx,
		'Action',
		[
			['filter-in-place', 'Filter the list, in-place'],
			['copy-to-location', 'Copy to another location'],
		],
		'filter-in-place',
	);
	const listRange = textInput(ctx, formatRange(region));
	const criteriaRange = textInput(ctx);
	const copyTo = textInput(ctx);
	const uniqueOnly = checkbox(ctx, 'Unique records only', false);

	const syncCopyTo = (): void => {
		copyTo.disabled = action.get() !== 'copy-to-location';
	};
	for (const input of action.inputs) input.addEventListener('change', syncCopyTo);
	syncCopyTo();

	return showDialog<AdvancedFilterResult>(ctx, {
		name: 'advanced-filter',
		heading: 'Advanced Filter',
		body: [
			action.element,
			field(ctx, 'List range:', listRange),
			field(ctx, 'Criteria range:', criteriaRange),
			field(ctx, 'Copy to:', copyTo),
			uniqueOnly.wrapper,
		],
		opened: () => listRange.focus(),
		submit: () => {
			const list = resolveReference(ctx, listRange.value);
			if (!list) return invalid(ctx, listRange, 'The reference is not valid.');
			const rawCriteria = criteriaRange.value.trim();
			const criteria = rawCriteria ? resolveReference(ctx, rawCriteria) : undefined;
			if (rawCriteria && !criteria)
				return invalid(ctx, criteriaRange, 'The reference is not valid.');
			const mode = action.get() as AdvancedFilterAction;
			const copy = mode === 'copy-to-location' ? resolveReference(ctx, copyTo.value) : undefined;
			if (mode === 'copy-to-location' && !copy)
				return invalid(ctx, copyTo, 'The reference is not valid.');
			const options: AdvancedFilterOptions = {
				listRange: list.range,
				action: mode,
				uniqueOnly: uniqueOnly.input.checked,
				...(criteria ? { criteriaRange: criteria.range, criteriaSheet: criteria.sheet } : {}),
				...(copy ? { copyTo: copy.range, copyToSheet: copy.sheet } : {}),
			};
			return t.session.applyAdvancedFilter(list.sheet, options);
		},
	});
}
