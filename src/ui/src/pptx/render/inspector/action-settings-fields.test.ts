import type { PptxSlide } from 'ooxml-core/pptx';
import { describe, expect, it } from 'vitest';

import { translationsEn } from '../../i18n/translations-en';
import type { ActionSettingsTarget } from './action-settings-fields';
import { ACTION_SETTINGS_SECTIONS } from './action-settings-fields';
import {
	coerceInspectorFieldValue,
	inspectorFieldPatch,
	resolveInspectorSections,
} from './inspector-field-model';
import { inspectorOptionLabel } from './inspector-option-label';

const slides = [
	{ id: 's1', rId: 'r1', slideNumber: 1, elements: [] },
	{ id: 's2', rId: 'r2', slideNumber: 2, elements: [] },
] as unknown as PptxSlide[];

const target = (extra: Partial<ActionSettingsTarget> = {}): ActionSettingsTarget => ({
	trigger: 'click',
	action: undefined,
	slides,
	customShows: [{ id: '7', name: 'Short version' }],
	...extra,
});

const patch = (t: ActionSettingsTarget, field: string, raw: string | boolean) =>
	inspectorFieldPatch(ACTION_SETTINGS_SECTIONS, 'actionSettings', field, t, raw);

describe('action settings fields', () => {
	it('holds a target-needing pick and commits a target-free one at once', () => {
		expect(patch(target(), 'type', 'url')).toStrictEqual({ pendingType: 'url' });
		expect(patch(target(), 'type', 'nextSlide')).toMatchObject({
			pendingType: 'nextSlide',
			commit: { trigger: 'click', type: 'nextSlide' },
		});
	});

	it('shows the target field of the pending type and commits once it is filled in', () => {
		const pending = target({ pendingType: 'url' });
		const [view] = resolveInspectorSections(ACTION_SETTINGS_SECTIONS, pending);
		expect(view!.fields.map((f) => f.id)).toStrictEqual(['type', 'urlUrl']);
		expect(patch(pending, 'urlUrl', 'https://example.com')).toMatchObject({
			commit: { type: 'url', url: 'https://example.com' },
		});
	});

	it('resolves the slide and custom show choices from the deck at render time', () => {
		const [view] = resolveInspectorSections(
			ACTION_SETTINGS_SECTIONS,
			target({ pendingType: 'slide' }),
		);
		const slide = view!.fields.find((f) => f.id === 'slide')!;
		expect(slide.options!.map((o) => [o.value, inspectorOptionLabel(o, (k) => k)])).toStrictEqual([
			['', 'pptx.action.slideNumberPlaceholder'],
			['0', '1'],
			['1', '2'],
		]);
		expect(patch(target({ pendingType: 'slide' }), 'slide', '1')).toMatchObject({
			commit: { type: 'slide', slideIndex: 1 },
		});
		// The "pick one" entry writes nothing.
		expect(patch(target({ pendingType: 'slide' }), 'slide', '')).toBeUndefined();
		const [shows] = resolveInspectorSections(
			ACTION_SETTINGS_SECTIONS,
			target({ pendingType: 'customShow' }),
		);
		expect(shows!.fields.find((f) => f.id === 'customShow')!.options!.at(-1)).toMatchObject({
			value: '7',
			label: 'Short version',
		});
	});

	it('bounds a number by a limit resolved from the target', () => {
		const field = {
			id: 'n',
			kind: 'number' as const,
			labelKey: 'k',
			maxOf: (t: { max: number }) => t.max,
			read: () => 0,
			write: (_t: { max: number }, v: unknown) => v,
		};
		expect(coerceInspectorFieldValue(field, '50', { max: 12 })).toBe(12);
	});

	it('labels every field with a dictionary key', () => {
		const keys = ACTION_SETTINGS_SECTIONS[0]!.fields.flatMap((f) => [
			f.labelKey,
			...(f.options ?? []).map((o) => o.labelKey),
		]);
		expect(keys.filter((key) => !(key in translationsEn))).toStrictEqual([]);
	});
});
